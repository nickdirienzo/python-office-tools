import magic
import os
import os.path
import xml.dom.minidom
import zipfile

class DocXReader(object):
	""" Reads document.xml to pull images from DOCX files.
		ONLY WORKS FOR MS WORD DOCX NOT OpenOffice """
	def __init__(self):
		self.temp_path = "/tmp/docx_extracts/"
		self.mime_finder = magic.Magic(mime=True)
	
	def find_images(self, docx_path):
		""" Finds the images and attempts to associate the name (without extension) to them,
			and it will return a dict such as {'path': 'name'} """
		filename = docx_path.split(".")[0]
		if not os.path.exists(self.temp_path):
			os.makedirs(self.temp_path)
		docx_zip = zipfile.ZipFile(docx_path)
		docx_zip.extractall(self.temp_path + filename)
		docx_zip.close()
		image_paths = self.find_image_paths(self.temp_path + filename)
		image_names = self.find_image_names(self.temp_path + filename)
		images = {}
		if len(image_paths) == len(image_names):
			for i in xrange(len(image_paths)):
				images[image_paths[i]] = image_names[i]
		else:
			for path in image_paths:
				images[path] = ""
		return images


	def find_image_paths(self, docx_extract_dir):
		""" Just for SW, we'll cheat a little and just look in word/media """
		image_paths = []
		media_dir = docx_extract_dir + "/word/media/"
		for media in os.listdir(media_dir):
			if self.mime_finder.from_file(media_dir + media) == "image/png":
				image_paths.append(media_dir + media)
		return image_paths

	def find_image_names(self, docx_extract_dir):
		""" Finding image names in the document.xml """
		image_names = []
		doc_xml = open(docx_extract_dir + "/word/document.xml")
		dom = xml.dom.minidom.parseString("".join(doc_xml.readlines()))
		doc_xml.close()
		body_node = dom.documentElement.childNodes[0]
		image_name_nodes = []
		nodes_to_visit = [body_node]
		while nodes_to_visit:
			node = nodes_to_visit.pop()
			for child in node.childNodes:
				nodes_to_visit.append(child)
			if node.__class__ == xml.dom.minidom.Element:
				if node.tagName == "pic:cNvPr":
					image_name_nodes.append(node.getAttribute("name").split(".")[0])
		return image_name_nodes
