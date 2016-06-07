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

Run ImageJ macro on P/D/I/SPW.
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
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders
from email.Utils import formatdate
import datetime

import omero.scripts as scripts
from omero.gateway import BlitzGateway
from omero.rtypes import *  # noqa

###############################################################################
# CONFIGURATION
###############################################################################

# The ImageJ classpath must contain the headless.jar before the ij.jar.
# The headless.jar with source code is available in the GitHub repository
# containing all the GDSC OMERO scripts.
IMAGEJ_CLASSPATH = "/Applications/Fiji.app/Contents/MacOS/ImageJ-macosx"
# IMAGEJ_CLASSPATH = "/Applications/ImageJ/headless.jar:/Applications/ImageJ/ImageJ64.app/Contents/Resources/Java/ij.jar:/Applications/ImageJ/plugins/bioformats_package.jar"

# The location of the ImageJ install (v1.45+). 
IMAGEJ_PATH = "/Applications/ImageJ"

# The e-mail address that messages are sent from. Make this a valid
# address so that the user can reply to the message.
ADMIN_EMAIL = 'b.ramalingam@dundee.ac.uk'


def get_original_file(conn, object_type, object_id, fileAnn_id=None):

    global tmp_dir
    if object_type == "Plate":
        omero_object = conn.getObject("Plate", int(object_id))
        if omero_object is None:
            sys.stderr.write("Error: Plate does not exist.\n")
            sys.exit(1)
    elif object_type == "Screen":
        omero_object = conn.getObject("Screen", int(object_id))
        if omero_object is None:
            sys.stderr.write("Error: Screen does not exist.\n")
            sys.exit(1)
    elif object_type == "Project":
        omero_object = conn.getObject("Project", int(object_id))
        if omero_object is None:
            sys.stderr.write("Error: Project does not exist.\n")
            sys.exit(1)
    elif object_type == "Dataset":
        omero_object = conn.getObject("Dataset", int(object_id))
        if omero_object is None:
            sys.stderr.write("Error: Dataset does not exist.\n")
            sys.exit(1)
    else:
        omero_object = conn.getObject("Image", int(object_id))
        if omero_object is None:
            sys.stderr.write("Error: Image does not exist.\n")
            sys.exit(1)
    fileAnn = None
    print "Listing files on %s %s..." % (object_type, object_id)
    for ann in omero_object.listAnnotations():
        if isinstance(ann, omero.gateway.FileAnnotationWrapper):
            fileName = ann.getFile().getName()
            print "   FileAnnotation ID:", ann.getId(), fileName,\
                "Size:", ann.getFile().getSize()
            # Pick file by Ann ID (or name if ID is None)
            if (fileAnn_id is None and fileName.endswith(".ijm")) or (
                    ann.getId() == fileAnn_id):
                fileAnn = ann
    if fileAnn is None:
        sys.stderr.write("Error: File does not exist.\n")
        sys.exit(1)
    print "Picked file annotation: %s %s" % (fileAnn.getId(),
                                             fileAnn.getFile().getName())

    file_path = os.path.join(tmp_dir, fileAnn.getFile().getName())

    f = open(str(file_path), 'w')
    print "\nDownloading file to", file_path, "..."
    try:
        for chunk in fileAnn.getFileInChunks():
            f.write(chunk)
    finally:
        f.close()

    return file_path


def extract_images(conn, images):
    """
    Extracts the images from OMERO.
    @param conn:   The BlitzGateway connection
    @param images: The list of images
    """
    global tmp_dir

    names = []
    tmp_dir = tempfile.mkdtemp(prefix='ImageDirectory')

    for img in images:
        if img is None:
            continue

        name = '%s/%s.ome.tif' % (tmp_dir, img.getId())
        e = conn.createExporter()
        e.addImage(img.getId())

        # Use a finally block to ensure clean-up of the exporter
        try:
            e.generateTiff()
            out = open(name, 'wb')

            read = 0
            while True:
                buf = e.read(read, 1000000)
                out.write(buf)
                if len(buf) < 1000000:
                    break
                read += len(buf)

            out.close()
        finally:
            e.close()

        names.append(name)

    print "LOG: Temp DIR : %s"% tmp_dir
    print "LOG: Done Exporting Images as ome-tiffs"
    return names

def upload_results(conn, results_path, params):
    """
    Uploads the results to each image as an annotation
    @param conn:         The BlitzGateway connection
    @param results:      Dict of (imageId,text_result) pairs
    @parms params:       The script parameters
    """
    global tmp_dir

    if not results_path:
        return

    results_csv = results_path[0];
    results_roi = results_path[1];

    current_time = datetime.datetime.now().time()
    result_csv_name = "imagej_csv_" + current_time.isoformat() + ".csv";
    result_roi_name = "imagej_roi_" + current_time.isoformat() + ".zip";
    
    object_ids = params["IDs"];
    object_id = object_ids[0];

    if params.get("Data_Type") == 'Dataset':
        dataset = conn.getObject("Dataset", int(object_id))
        name = "%d_%s" % (int(object_id), result_csv_name)
        ann = conn.createFileAnnfromLocalFile(
            results_csv, origFilePathAndName=name,
            ns='ImageJ_csv')
        print "Attaching FileAnnotation to Dataset: ", "File ID:", ann.getId(), \
            ",", ann.getFile().getName(), "Size:", ann.getFile().getSize()
        dataset.linkAnnotation(ann)

        name1 = "%d_%s" % (int(object_id), result_roi_name)
        ann1 = conn.createFileAnnfromLocalFile(
            results_roi, origFilePathAndName=name1,
            ns='ImageJ_roi_zip')
        print "Attaching FileAnnotation to Dataset: ", "File ID:", ann1.getId(), \
            ",", ann1.getFile().getName(), "Size:", ann1.getFile().getSize()
        dataset.linkAnnotation(ann1)

        print "Results attached to Dataset"

    else:
        return      

def run(conn, params):
    """
    For each image defined in the script parameters run the correlation
    analyser and load the result into OMERO.
    Returns the number of images processed or (-1) if there is a
    parameter error).
    @param conn:   The BlitzGateway connection
    @param params: The script parameters
    """
    global tmp_dir

    print "Parameters = %s" % params

    if not params.get("IDs"):
        return -1

    images = []
    if params.get("Data_Type") == 'Image':
        objects = conn.getObjects("Image", params["IDs"])
        images = list(objects)
    else:
        for dsId in params["IDs"]:
            ds = conn.getObject("Dataset", dsId)
            if ds:
                for i in ds.listChildren():
                    images.append(i)

    # Extract images
    image_names = extract_images(conn, images)
    print "LOG: Done printing %d images" % len(image_names)
    #Point at the File Annotation
    #macro_file = get_original_file(conn, object_type, object_id, fileAnn_id=None)
    object_ids = params["IDs"]
    object_id = object_ids[0]
    fileAnn_id = None
    if "File_Annotation" in params:
        fileAnn_id = long(params["File_Annotation"])
    dataType = params["Data_Type"]
    macro_file = get_original_file(conn, dataType, object_id, fileAnn_id)

    # Run ImageJ
    results = run_imagej(conn, image_names, macro_file)

    if results:
        # Upload results
        upload_results(conn, results, params)
    else:
        print "ERROR: No results generated"
    
    #Delete temp files
    try:
        for name in glob.glob("%s/*" % tmp_dir):
            os.remove(name)
    except:
        pass
    os.rmdir(tmp_dir)

    return results


def run_imagej(conn, image_names, macro_file=None):

    global tmp_dir

    if not image_names:
        return results

    with open(macro_file,'r') as myfile:
        data = myfile.read()

    macro_open_file = os.path.join(tmp_dir, "open_file.ijm")

    out = open(macro_open_file, 'wb')
    
    #read file annotation macro_file and append to the macro_open_file
    csvPath = os.path.join(tmp_dir,"imageJresults.csv");
    roiPath = os.path.join(tmp_dir,"imageJrois.zip");
    initialString = """// Stack analysis of omero images macro
            setBatchMode(true);
            run("Bio-Formats Macro Extensions");"""
    for i, name in enumerate(image_names):
        if i==0:
            header = initialString;
        else:
            header = "";
        out.write("""%s
            imps = Ext.openImagePlus("%s");
        """ % (header, name))
        out.write(data)
    out.write("""saveAs("Results", "%s");
roiManager("Save","%s");
run("Quit");
        """ % (csvPath, roiPath)) 
    out.close()
    # Run ImageJ
    try:
        args = [IMAGEJ_CLASSPATH, "-macro",
                macro_open_file]

        # debug
        cmd = " ".join(args)
        print "Script command = %s" % cmd

        # Run the command
        results = subprocess.Popen(args, stdout=subprocess.PIPE, stdin=subprocess.PIPE).communicate()
        std_out = results[0]
        std_err = results[1]
        print std_out
        print "Done running ImageJ macro"
    
    except OSError, e:
        print >>sys.stderr, "Execution failed:", e    

    results_path = ["csvPath","roiPath"]
    results_path[0] = csvPath;
    results_path[1] = roiPath;
    #return an array of file paths (results[0]=csv,results[1]=rois)
    return results_path


if __name__ == "__main__":
    dataTypes = [rstring('Project'), rstring('Dataset'), rstring('Image'), rstring('Plate'), rstring('Screen')]
    client = scripts.client(
        'Batch_Analysis_IJMacro.py',
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

        scripts.String(
            "File_Annotation", grouping="3",
            description="File ID containing ImageJ macro (extension:*.ijm)."),

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
        results = run(conn, scriptParams)
        message = "Done"
        client.setOutput("Message", rstring(message))

    finally:
        client.closeSession()