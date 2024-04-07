from zipfile import ZipFile
import os 
import xml.etree.ElementTree as ET
from io import BytesIO
from pprint import pprint


def main():
    doc = Docx("test.docx")
    doc.process()
    print(doc.metadata)
    print(doc.data)

class Document:
    def __init__(self, filepath, ftype):
        self.filepath = filepath
        self.ftype = ftype.lower()
        self.data = {}
        self.metadata = {}

    def process_data(self):
        raise NotImplementedError

    def process_metadata(self):
        raise NotImplementedError
    
    def run_enrichments(self):
        raise NotImplementedError
    
    def process(self):
        self.process_data()
        self.process_metadata()
        self.run_enrichments()


class Docx(Document):
    def __init__(self, filepath):
        Document.__init__(self, filepath, "docx")
    
    def process_data(self):
        with open(self.filepath, "rb") as fh:
            _zip = ZipFile(fh)
            target_files = _zip.namelist()
            target_files = [target_file for target_file in target_files if 'word/' in target_file]
            for target_file in target_files:
                document_xml_content = _zip.read(target_file)
                self.process_data_xml(document_xml_content)
            _zip.close()
        

    def process_data_xml(self, xml_content):
        tree = ET.XML(xml_content)
        for node in tree.iter():
            if not node.text:
                continue
            tag = node.tag.split('}')[-1]
            if not tag in self.data.keys():
                self.data[tag] = node.text
            else:
                self.data[tag] += "\n" + node.text


    def process_metadata(self):
        with open(self.filepath, "rb") as fh:
            _zip = ZipFile(fh)
            app_xml_content = _zip.read("docProps/app.xml")
            core_xml_content = _zip.read("docProps/core.xml")
            _zip.close()
        
        self.process_metadata_xml(app_xml_content)
        self.process_metadata_xml(core_xml_content)


    def process_metadata_xml(self, xml_content):
        tree = ET.XML(xml_content)
        for node in tree.iter():
            if not node.text:
                continue
            tag = node.tag.split('}')[-1]
            if not node.tag in self.metadata.keys():
                self.metadata[tag] = [node.text]
            else:
                self.metadata[tag] += [node.text]


    def run_enrichments(self):
        None
    

if __name__ == "__main__":
    main()
