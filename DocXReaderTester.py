if __name__ == "__main__":
	import AssetReader
	docx_reader = AssetReader.DocXReader()
	print docx_reader.find_images("GestaltBoxDemo.docx")
