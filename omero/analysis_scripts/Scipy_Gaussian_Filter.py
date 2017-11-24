# coding=utf-8
'''
-----------------------------------------------------------------------------
  Copyright (C) 2014 Glencoe Software, Inc. All rights reserved.


  This program is free software; you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation; either version 2 of the License, or
  (at your option) any later version.
  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License along
  with this program; if not, write to the Free Software Foundation, Inc.,
  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

------------------------------------------------------------------------------

Run Gaussian Filter from Scipy on P/D/I/SPW.
'''

import omero
from omero.model import PlateI, ScreenI

import os
import sys
import subprocess
import re
import tempfile
import platform
import glob
import smtplib
# For guessing MIME type based on file name extension
import mimetypes
import datetime

import omero
import omero.scripts as scripts
from omero.rtypes import wrap, rlong
from omero.gateway import OriginalFileWrapper
from omero.gateway import BlitzGateway
from omero.rtypes import *  # noqa
from omeroweb.webgateway.marshal import imageMarshal

import json
import ConfigParser
from cStringIO import StringIO

import scipy
import scipy.ndimage as spi


def run(conn, params):
    """
    For each image defined in the script parameters run the correlation
    analyser and load the result into OMERO.
    Returns the number of images processed or (-1) if there is a
    parameter error).
    @param conn:   The BlitzGateway connection
    @param params: The script parameters
    """

    print "Parameters = %s" % params

    if not params.get("Kernel_Window_Size"):
        print "Please enter a valid window size for the Gaussian kernel"
        return -1
    window_size = params.get("Kernel_Window_Size")
    
    if not params.get("Sigma"):
        print "Please enter a valid value for Sigma (>0)"
        return -1
    sigma = params.get("Sigma")

    if not params.get("IDs"):
        print "Please enter a valid omero_object (Project/Dataset/Image) Id"
        return -1

    images = []

    if params.get("Data_Type") == 'Dataset':
        for dsId in params["IDs"]:
            dataset = conn.getObject("Dataset", dsId)
            if dataset:
                for image in dataset.listChildren():
                    images.append(image)

    # Extract images
    image_ids = []
    for image in images:

        sizeZ = image.getSizeZ()
        sizeC = image.getSizeC()
        sizeT = image.getSizeT()
        zctList = []
        for z in range(sizeZ):
            for c in range(sizeC):
                for t in range(sizeT):
                    zctList.append( (z,c,t) ) 


        i = conn.createImageFromNumpySeq(planeGen(image, zctList, window_size, sigma), image.getName() + "_numpy_image", sizeZ, sizeC, sizeT, description="Gaussian Filter",dataset=dataset)
        print 'Created new Image:%s Name:"%s"' % (i.getId(), i.getName())
        imageId = i.getId()
        image_ids.append(imageId)

    print "LOG: Done printing %d images" % len(images)

    return image_ids

def planeGen(image, zctList, window, sigma):
    """generator will yield planes"""

    planes = image.getPrimaryPixels().getPlanes(zctList)
    for p in planes:
        # data = rolling_ball_background(p, radius, light_background=True, smoothing=True)
        data = spi.filters.gaussian_filter(p, sigma=sigma, truncate=(((window - 1)/2)-0.5)/sigma)
        yield data

### OMERO Figure Methods


def get_panel_json(image, x, y, width, height, channel=None):

    rv = imageMarshal(image)

    if channel is not None:
        for idx, ch in enumerate(rv['channels']):
            ch['active'] = idx == channel
            # print ch['color']
            ch['color'] = 'ffffff'

    img_json = {
        "labels":[],
        "height": height,
        "channels": rv['channels'],
        "width": width,
        "sizeT": rv['size']['t'],
        "sizeZ": rv['size']['z'],
        "dx":0,
        "dy":0,
        "rotation":0,
        "imageId":image.getId(),
        "name":"438CTR_01_5_R3D_D3D.dv",
        "orig_width": rv['size']['width'],
        "zoom":100,
        "shapes":[],
        "orig_height": rv['size']['height'],
        "theZ": rv['rdefs']['defaultZ'],
        "y": y,
        "x": x,
        "theT": rv['rdefs']['defaultT']
    }
    return img_json


def get_labels_json(panel_json, column, row):

    labels = []

    channels = panel_json['channels']
    imagename = panel_json['name']
    if row == 0:
        labels.append({"text":channels[column]['label'],
                       "size":4,
                       "position":"leftvert",
                       "color":"000000"
                     })
    if column == 0:
        labels.append({"text": imagename,
                       "size":4,
                       "position":"top",
                       "color":"000000"
                     })
    return labels

def create_figure_file(conn, image_ids):

    width = 512/10;
    height = 512/10;
    spacing_x = 512/50;
    spacing_y = 512/50;
    page_width = (width + spacing_x) * (5) * 1.25
    page_height = (height + spacing_y) * (len(image_ids)) * 1.25

    JSON_FILEANN_NS = "omero.web.figure.json"

    figure_json = {"version":2,
                   "paper_width":page_width,
                   "paper_height":page_height,
                   "page_size":"mm",
                   "figureName":"from script",
                   }

    curr_x = 0
    curr_y = 0
    panels_json = []
    column_count = 2
    offset = 10

    for z, image_id in enumerate(image_ids):
        image = conn.getObject('Image', image_id)
        curr_x = z * (width + spacing_x) + offset
        for c in range(image.getSizeC()):
            curr_y = c * (height + spacing_y) + offset
            j = get_panel_json(image, curr_x, curr_y, width, height, c)
            j['labels'] = get_labels_json(j, c, z)
            panels_json.append(j)

    figure_json['panels'] = panels_json

    figure_name = figure_json['figureName']
    if len(figure_json['panels']) == 0:
        raise Exception('No Panels')
    first_img_id = figure_json['panels'][0]['imageId']

    # we store json in description field...
    description = {}
    description['name'] = figure_name
    description['imageId'] = first_img_id

    # Try to set Group context to the same as first image
    conn.SERVICE_OPTS.setOmeroGroup('-1')
    i = conn.getObject("Image", first_img_id)
    gid = i.getDetails().getGroup().getId()
    conn.SERVICE_OPTS.setOmeroGroup(gid)

    json_string = json.dumps(figure_json)
    file_size = len(json_string)
    f = StringIO()
    # f.write(figure_json)
    json.dump(figure_json, f)

    update = conn.getUpdateService()
    orig_file = create_original_file_from_file_obj(conn, 
        f, '', figure_name, file_size, mimetype="application/json")
    fa = omero.model.FileAnnotationI()
    fa.setFile(omero.model.OriginalFileI(orig_file.getId(), False))
    fa.setNs(wrap(JSON_FILEANN_NS))
    desc = json.dumps(description)
    fa.setDescription(wrap(desc))
    fa = update.saveAndReturnObject(fa, conn.SERVICE_OPTS)
    file_id = fa.getId().getValue()

    print "Figure Created:", file_id


def create_original_file_from_file_obj(conn, fo, path, name, file_size, mimetype=None):
    """
    Creates and saves an OMERO OriginalFile from a file object
    """
    raw_file_store = conn.createRawFileStore()

    # create original file, set name, path, mimetype
    original_file = omero.model.OriginalFileI()
    original_file.setName(wrap(name))
    original_file.setPath(wrap(path))
    if mimetype:
        original_file.mimetype = wrap(mimetype)
    original_file.setSize(rlong(file_size))
    # set sha1  # ONLY for OMERO-4
    try:
        import hashlib
        hash_sha1 = hashlib.sha1
    except:
        import sha
        hash_sha1 = sha.new
    try:
        fo.seek(0)
        h = hash_sha1()
        h.update(fo.read())
        original_file.setSha1(wrap(h.hexdigest()))
    except:
        pass       # OMERO-5 doesn't need this
    upd = conn.getUpdateService()
    original_file = upd.saveAndReturnObject(original_file, conn.SERVICE_OPTS)

    # upload file
    fo.seek(0)
    raw_file_store.setFileId(original_file.getId().getValue(),
                             conn.SERVICE_OPTS)
    buf = 10000
    for pos in range(0, long(file_size), buf):
        block = None
        if file_size-pos < buf:
            block_size = file_size-pos
        else:
            block_size = buf
        fo.seek(pos)
        block = fo.read(block_size)
        raw_file_store.write(block, pos, block_size, conn.SERVICE_OPTS)
    # https://github.com/openmicroscopy/openmicroscopy/pull/2006
    original_file = raw_file_store.save(conn.SERVICE_OPTS)
    raw_file_store.close()
    return OriginalFileWrapper(conn, original_file)


if __name__ == "__main__":
    dataTypes = [rstring('Dataset')]
    client = scripts.client(
        'Scipy_Gaussian_Filter.py',
        """
    This script processes a ijm file, attached to a P/D/I/Screen or Plate,
        """,
        scripts.String(
            "Data_Type", optional=False, grouping="1",
            description="Choose source of images",
            values=dataTypes, default="Dataset"),

        scripts.List(
            "IDs", optional=False, grouping="2",
            description="Project or Dataset or Image or Plate or Screen ID.").ofType(rlong(0)),

        scripts.Int(
            "Kernel_Window_Size", optional=False, grouping="3", default=20,
            description="Window size for the gaussian filter"),

        scripts.Int(
            "Sigma", optional=False, grouping="4", default=2,
            description="Sigma for the gaussian filter"),

        scripts.Bool(
            "Create_Omero_Figure", default=True, grouping="5",
            description="Create An OMERO.Figure from the resultant images"),
        
        authors=["Balaji Ramalingam", "OME Team"],
        institutions=["University of Dundee"],
        contact="ome-users@lists.openmicroscopy.org.uk",
    )

    try:
        # process the list of args above.
        scriptParams = {}
        for key in client.getInputKeys():
            if client.getInput(key):
                scriptParams[key] = client.getInput(key, unwrap=True)
        print scriptParams

        # wrap client to use the Blitz Gateway
        conn = BlitzGateway(client_obj=client)
        # # Call the main script - returns the number of images processed
        image_ids = run(conn, scriptParams)
        
        if scriptParams["Create_Omero_Figure"]:
            create_figure_file(conn, image_ids)

        message = "Done"
        client.setOutput("Message", rstring(message))

    finally:
        client.closeSession()