# @ File (label = "directory to concat", style = "directory") indir
# @ File (label = "directory for output", style = "directory") outdir
# @ Boolean fixt0
# @ Boolean (label = "regex for merging several image types") well_position

'''
To run concat_.py in fiji you have two options
Make it a plugin (have to be done only once):
1) copy concat_.py in the Fiji.app plugins
2) Open Fiji, Help -> Refresh Menus
The macro will appear in the plugin menu as concat

Run it from the macro editor (have to be done everytime you use the macro):
1) Plugins -> New -> Macro
2) Open concat_.py in the macro editor
3) push run
Author Antonio Politi, MPIBPC
Modified May 2021
'''

import traceback
import os
import re
from time import time
import datetime
from glob import glob
from ij import IJ, ImagePlus, ImageStack
#Bioformats specific stuff
from loci.plugins import BF
# the not showing in the import is a bug
import loci.plugins.in.ImporterOptions as ImporterOptions
#import ImporterOptions
from loci.formats import ImageReader
from loci.formats import MetadataTools
from loci.formats import ImageWriter
from loci.formats.tiff import TiffParser
from loci.formats.tiff import TiffSaver
from ome.xml.model.enums import DimensionOrder
from ome.xml.model.enums import PixelType
from ome.xml.model.primitives import PositiveInteger, NonNegativeInteger
from ome.units.quantity import Time
from loci.common import RandomAccessInputStream
from loci.common import RandomAccessOutputStream

pattern =  ['(\S+\d+|\d+\S+)(_+T|_+t)(\d+)\.(lsm$|czi$)', '(\S+\d+|\d+\S+)(_+T|_+t)(\d+)_Out\.(lsm$|czi$)']

def run_onfiles(infiles = None, outdir = None):
	IJ.run("Close All")
	IJ.log(" ")
	IJ.log("!!concat: concatenation of lsm, czi created with AutofocusScreen VBA macro!!")
	if outdir is None:
		outdir = IJ.getDirectory("Choose output directory. If cancel macro uses respective local directories")
	if outdir == '': 
		outdir = None

	if outdir is not None:
		if not os.path.exists(outdir):
			os.mkdir(outdir)

	process_files_ome(infiles, outdir)

def run(indir=None,outdir=None):
	IJ.run("Close All")
	IJ.log(" ")
	IJ.log("!!concat: concatenation of lsm, czi created with AutofocusScreen VBA macro!!")
	'''Main execution. Ask users for input and output directories if not specified'''
	if indir is None:
		indir = IJ.getDirectory("Choose the input directory")
	if indir is None or indir == '':
		IJ.showMessage("Concat: No input directory defined!")
		return
	if outdir is None:
		outdir = IJ.getDirectory("Choose output directory. If cancel macro uses respective local directories")
	if outdir == '': 
		outdir = None



	if outdir is not None:
		if not os.path.exists(outdir):
			os.mkdir(outdir)

	start_time = time()
	for root, dirs, files in os.walk(indir, topdown = False):
		locfiles = glob(root+"/*.lsm")
		locfiles= locfiles + glob(root+"/*.czi")
		if locfiles is not None:
			find_timepoints(root, locfiles, outdir)
	return start_time

def getFilesToProcess(indir):
	"""get files to process according to a pattern"""
	pattern = ".+_(?P<jobidx>\d+)_W(?P<well>\d+)_P(?P<position>\d+)(_T(?P<timepoint>\d+))?.(ome.tif|lsm|czi)"
	files = os.listdir(indir)
	file_with_indexes = list() 
	for afile in files:
		m = re.match(pattern, afile)
		if m is not None:
			tpoint = m.group('timepoint')
			if tpoint is not None:
				tpoint = int(m.group('timepoint'))
			file_with_indexes.append([afile,  int(m.group('jobidx')), int(m.group('well')), int(m.group('position')), tpoint])
	
	# find matching well and position
	posused = list()
	wellused = list()
	concat_files = list()
	for afile in file_with_indexes:	
		well = afile[2]
		pos = afile[3]
		if well in wellused and pos in posused:
			continue
		concat_files.append([os.path.join(indir, x[0]) for x in file_with_indexes if x[2] == well and x[3] == pos])
		posused.append(pos)
		wellused.append(well)
	
	return concat_files

def find_timepoints(root, files, outdir):
	'''Find files with same root and process them'''
	while len(files) > 0:
		for file in files:
			filename =  os.path.basename(file)

			for patt  in pattern:
				result = re.match(patt, filename)
				if result is not None:
					break
			if result is None:
				files.remove(file)
				continue
			baseName = result.group(1)
			locfiles = glob(os.path.join(root, baseName + result.group(2)+'*'+result.group(4)))
			print(locfiles)
			if locfiles is not None:
				for i in range(0,len(locfiles)):
					try:
						files.remove(locfiles[i])
					except:
						pass
				process_time_points(root, locfiles, outdir)

def maxSizeT(files):
	'''Maximal number of time points when concatenating files with each some time-lapse data'''
	options = ImporterOptions()
	sizeT = 0
	for fileName in files:
		options.setId(fileName)
		options.setVirtual(1)
		image = BF.openImagePlus(options)
		image = image[0]
		sizeT_file = image.getNFrames()
		sizeT = sizeT + sizeT_file
		image.close()
	return sizeT

def process_files_ome( files, outdir):
	''' Concatenate ome.tiff files. From different settings Jobs '''

	options = ImporterOptions()
	# compute maximal number of frames

	reader = ImageReader()

	options.setId(files[0])
	options.setVirtual(1)
	image = BF.openImagePlus(options)
	image = image[0]
	sizeT = maxSizeT(files)
	
	# Create some default ome dump file from first file
	reader.setMetadataStore(MetadataTools.createOMEXMLMetadata())
	reader.setId(files[0])
	omeOut = reader.getMetadataStore()
	omeOut = setUpXml(omeOut, image, sizeT)
	nrplanes_per_timepoint = omeOut.getPixelsSizeC(0).getValue()*omeOut.getPixelsSizeZ(0).getValue()
	reader.close()

	images_total = 0
	itime = 0
	itime_local = 0

	outName = re.match('(\S+\d+|\d+\S+)(_+T|_+t)(\d+)\.(lsm$|czi$|ome.tif$)', os.path.basename(files[0]))

	outfile =  os.path.join(outdir, outName.group(1) + '_final.ome.tif')

	print(outfile)

	for ifile, fileName in enumerate(files):

		omeMeta = MetadataTools.createOMEXMLMetadata()
		reader.setMetadataStore(omeMeta)
		reader.setId(fileName)
		nrImages = reader.getImageCount()
		if ifile == 0:
			T0 = omeMeta.getPlaneDeltaT(0,0).value()
			unit =  omeMeta.getPlaneDeltaT(0,0).unit()
		for i in range(0, nrImages):
			
			itime_local = itime + (i/nrplanes_per_timepoint)
			dT = omeMeta.getPlaneDeltaT(0,i).value() - T0
			omeOut.setPlaneDeltaT(Time(dT, unit),0, i + images_total)
			omeOut.setPlanePositionX(omeMeta.getPlanePositionX(0,i), 0, i + images_total)
			omeOut.setPlanePositionY(omeMeta.getPlanePositionY(0,i), 0, i + images_total)
			omeOut.setPlanePositionZ(omeMeta.getPlanePositionZ(0,i), 0, i + images_total)
			omeOut.setPlaneTheC(omeMeta.getPlaneTheC(0,i), 0, i + images_total)
			omeOut.setPlaneTheT(NonNegativeInteger(itime_local), 0, i + images_total)
			omeOut.setPlaneTheZ(omeMeta.getPlaneTheZ(0,i), 0, i + images_total)
		itime = itime_local  + 1
		images_total = images_total + nrImages
		reader.close()

	try:
		incr = float(dT/(sizeT - 1))
	except:
		incr = 0

	try:
		omeOut.setPixelsTimeIncrement(incr, 0)
	except TypeError:
		#new Bioformats >5.1.x
		omeOut.setPixelsTimeIncrement(Time(incr, unit),0)

	outfile = concatenateImagePlus(files, sizeT, outfile)
	if outfile is not None:
		filein = RandomAccessInputStream(outfile)
		fileout = RandomAccessOutputStream(outfile)
		saver = TiffSaver(fileout, outfile)
		saver.overwriteComment(filein,omeOut.dumpXML())
		fileout.close()
		filein.close()


def process_time_points(root, files, outdir):
	'''Concatenate images and write ome.tiff file. If image contains already multiple time points just copy the image'''
	concat = 1
	files.sort()
	options = ImporterOptions()
	options.setId(files[0])
	options.setVirtual(1)
	image = BF.openImagePlus(options)
	image = image[0]
	if image.getNFrames() > 1:
		basename = os.path.basename(files[0])
		basename = os.path.splitext(basename)[0]
		IJ.log(files[0] + " Contains multiple time points. Can only concatenate single time points! Just export as ome.tif")
		IJ.run(image, "Bio-Formats Exporter", "save="+ os.path.join(outdir, basename + ".ome.tif") +" export compression=Uncompressed");
		image.close()
		return
	
	width  = image.getWidth()
	height = image.getHeight()
	for patt in pattern:
		outName = re.match(patt, os.path.basename(files[0]))
		if outName is None:
			continue
		if outdir is None:
			outfile = os.path.join(root, outName.group(1) + '.ome.tif')
		else:
			outfile =  os.path.join(outdir, outName.group(1) + '.ome.tif')
		reader = ImageReader()
		reader.setMetadataStore(MetadataTools.createOMEXMLMetadata())
		reader.setId(files[0])
		timeInfo = []
		omeOut = reader.getMetadataStore()
		omeOut = setUpXml(omeOut, image, len(files))
		reader.close()
		image.close()
		IJ.log ('Concatenates ' + os.path.join(root, outName.group(1) + '.ome.tif'))
		itime = 0
		try:
			for ifile, fileName in enumerate(files):
				omeMeta = MetadataTools.createOMEXMLMetadata()
	
				reader.setMetadataStore(omeMeta)
				reader.setId(fileName)
				#print omeMeta.getPlaneDeltaT(0,0)
				#print omeMeta.getPixelsTimeIncrement(0)
				
				if fileName.endswith('.czi'):
					if ifile == 0:
						T0 = omeMeta.getPlaneDeltaT(0,0).value()
					if fixt0:
						T0 = 0
					dT = omeMeta.getPlaneDeltaT(0,0).value() - T0
					unit =  omeMeta.getPlaneDeltaT(0,0).unit()
				else:
					timeInfo.append(getTimePoint(reader, omeMeta))
	 				unit = omeMeta.getPixelsTimeIncrement(0).unit()
					try:
						dT = round(timeInfo[files.index(fileName)]-timeInfo[0],2)
					except:
						dT = (timeInfo[files.index(fileName)]-timeInfo[0]).seconds
				
				nrImages = reader.getImageCount()
				for i in range(0, reader.getImageCount()):
					try:
						omeOut.setPlaneDeltaT(dT, 0, i + itime*nrImages)
					except TypeError:
						omeOut.setPlaneDeltaT(Time(dT, unit),0, i + itime*nrImages)
					omeOut.setPlanePositionX(omeOut.getPlanePositionX(0,i), 0, i + itime*nrImages)
					omeOut.setPlanePositionY(omeOut.getPlanePositionY(0,i), 0, i + itime*nrImages)
					omeOut.setPlanePositionZ(omeOut.getPlanePositionZ(0,i), 0, i + itime*nrImages)
					omeOut.setPlaneTheC(omeOut.getPlaneTheC(0,i), 0, i + itime*nrImages)
					omeOut.setPlaneTheT(NonNegativeInteger(itime), 0, i + itime*nrImages)
					omeOut.setPlaneTheZ(omeOut.getPlaneTheZ(0,i), 0, i + itime*nrImages)
				itime = itime + 1
				reader.close()
	
				IJ.showProgress(files.index(fileName), len(files))
			try:
				incr = float(dT/(len(files)-1))
			except:
				incr = 0
			
			try:
				omeOut.setPixelsTimeIncrement(incr, 0)
			except TypeError:
				#new Bioformats >5.1.x
				omeOut.setPixelsTimeIncrement(Time(incr, unit),0)
			
			outfile = concatenateImagePlus(files, len(files),  outfile)
			if outfile is not None:
				filein = RandomAccessInputStream(outfile)
				fileout = RandomAccessOutputStream(outfile)
				saver = TiffSaver(fileout, outfile)
				saver.overwriteComment(filein,omeOut.dumpXML())
				fileout.close()
				filein.close()
	
	
		except:
			traceback.print_exc()
		finally:
			#close all possible open files
			try:
				reader.close()
			except:
				pass
			try:
				filein.close()
			except:
				pass
			try:
				fileout.close()
			except:
				pass

def getTimePoint(reader, omeMeta):
	""" Extract timeStamp from file """
	time = datetime.datetime.strptime(str(omeMeta.getImageAcquisitionDate(0)), "%Y-%m-%dT%H:%M:%S")
	time = [reader.getSeriesMetadataValue("TimeStamp0"), reader.getSeriesMetadataValue("TimeStamp #1"), omeMeta.getPlaneDeltaT(0,0)]
	time = [x for x in time if x is not None]
	return time[0]



def setUpXml(ome, image, sizeT):
	"""setup Xml standard for concatenated file"""
	ome.setImageID("Image:0", 0)
	ome.setPixelsID("Pixels:0", 0)
	ome.setPixelsDimensionOrder(DimensionOrder.XYCZT,0)

	if image.getBitDepth() == 8:
		pixels = PixelType.UINT8
	if image.getBitDepth() == 12:
		pixels = PixelType.UINT12
	if image.getBitDepth() == 16:
		pixels = PixelType.UINT16

	ome.setPixelsType(pixels, 0)
	ome.setPixelsSizeX(PositiveInteger(image.getWidth()), 0)
	ome.setPixelsSizeY(PositiveInteger(image.getHeight()), 0)
	ome.setPixelsSizeZ(PositiveInteger(image.getNSlices()), 0)
	ome.setPixelsSizeT(PositiveInteger(sizeT), 0)
	ome.setPixelsSizeC(PositiveInteger(image.getNChannels()), 0)
	return ome


def addPositionName(position, outfile):
	"""parse outfile and add a position information at the end"""
	filename, extension = os.path.splitext(outfile)
	return filename + '_P' + str(position) + extension

def concatenateImagePlus(files, sizeT, outfile):
	"""concatenate images contained in files and save in outfile"""
	'''
	if len(files) == 1:
		IJ.log(files[0] + " has only one time point! Nothing to concatenate!")
		return
	'''
	options = ImporterOptions()
	try:
		filestr = files[0].getPath()
	except AttributeError:
		filestr = files[0]
	options.setId(filestr)
	options.setVirtual(1)
	options.setOpenAllSeries(1)
	options.setQuiet(1)
	images = BF.openImagePlus(options)
	imageG = images[0]
	nrPositions = len(images)
	options.setOpenAllSeries(0)

	for i in range(0, nrPositions):
		concatImgPlus = IJ.createHyperStack("ConcatFile", imageG.getWidth(), imageG.getHeight(), imageG.getNChannels(), imageG.getNSlices(), sizeT, imageG.getBitDepth())
		concatStack = ImageStack(imageG.getWidth(), imageG.getHeight())
		IJ.showStatus("Concatenating files")
		for file in files:
			try:
				filestr = file.getPath()
			except AttributeError:
				filestr = file
			try:
				IJ.log("	Add file " + filestr)
				options.setSeriesOn(i,1)
				options.setId(filestr)
				image = BF.openImagePlus(options)[0]
				imageStack = image.getImageStack()
				sliceNr = imageStack.getSize()
				for j in range(1, sliceNr+1):
					concatStack.addSlice(imageStack.getProcessor(j))
				image.close()
				options.setSeriesOn(i,0)
			except:
				traceback.print_exc()
				IJ.log(filestr + " failed to concatenate!")
			IJ.showProgress(files.index(file), len(files))
		concatImgPlus.setStack(concatStack)
		concatImgPlus.setCalibration(image.getCalibration())
		if len(images) > 1:
			outfileP = addPositionName(i+1,outfile)
			IJ.saveAs(concatImgPlus, "Tiff",  outfileP)
		else:
			IJ.saveAs(concatImgPlus, "Tiff",  outfile)
		concatImgPlus.close()
	return outfile

if not well_position:
	run(indir.getPath(), outdir.getPath())
else:
	recursive = 1
	if recursive:
		maindir = indir.getPath()
		dirs = os.listdir(maindir)
		for adir in dirs:
			locdir = os.path.join(maindir,adir)
			if os.path.isdir(os.path.join(maindir,adir)):
				concat_files = getFilesToProcess(locdir)
				for cfiles in concat_files:
					run_onfiles(cfiles,  outdir.getPath())			
	else:
		concat_files = getFilesToProcess(indir.getPath())
		print(concat_files)
		for cfiles in concat_files:
			run_onfiles(cfiles,  outdir.getPath())
	
#if __name__=="__main__":
#	run()
#dC = directoryChooser()
