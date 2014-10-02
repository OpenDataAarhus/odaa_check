__author__ = 'arne'
from pdfminer.pdfparser import PDFParser
from pdfminer.psparser import PSLiteral
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdftypes import PDFObjRef
from pdfminer.layout import LAParams, LTTextBoxHorizontal
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfdocument import PDFDocument, PDFTextExtractionNotAllowed
from pdfminer.pdfpage import PDFPage
from collections import defaultdict, namedtuple

TextBlock= namedtuple("TextBlock", ["x", "y", "text"])

class Parser( object ):
    """Parse the PDF.

    1.  Get the annotations into the self.fields dictionary.

    2.  Get the text into a dictionary of text blocks.
        The key to the dictionary is page number (1-based).
        The value in the dictionary is a sequence of items in (-y, x) order.
        That is approximately top-to-bottom, left-to-right.
    """
    def __init__( self ):
        self.fields = {}
        self.text= {}

    def load( self, open_file ):
        self.fields = {}
        self.text= {}

        # Create a PDF parser object associated with the file object.
        parser = PDFParser(open_file)
        # Create a PDF document object that stores the document structure.
        doc = PDFDocument(parser)
        # Connect the parser and document objects.
        parser.set_document(doc)
        #doc.set_parser(parser)
        # Supply the password for initialization.
        # (If no password is set, give an empty string.)
        #doc.initialize('')
        # Check if the document allows text extraction. If not, abort.
        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed
        # Create a PDF resource manager object that stores shared resources.
        rsrcmgr = PDFResourceManager()
        # Set parameters for analysis.
        laparams = LAParams()
        # Create a PDF page aggregator object.
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # Create a PDF interpreter object.
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        # Process each page contained in the document.
        pages = PDFPage(doc, 0, "Null")


        for pgnum, page in enumerate( pages.get_pages() ):
            interpreter.process_page(page)
            if page.annots:
                self._build_annotations( page )
            txt= self._get_text( device )
            self.text[pgnum+1]= txt

    def _build_annotations( self, page ):
        for annot in page.annots.resolve():
            if isinstance( annot, PDFObjRef ):
                annot= annot.resolve()
                assert annot['Type'].name == "Annot", repr(annot)
                if annot['Subtype'].name == "Widget":
                    if annot['FT'].name == "Btn":
                        assert annot['T'] not in self.fields
                        self.fields[ annot['T'] ] = annot['V'].name
                    elif annot['FT'].name == "Tx":
                        assert annot['T'] not in self.fields
                        self.fields[ annot['T'] ] = annot['V']
                    elif annot['FT'].name == "Ch":
                        assert annot['T'] not in self.fields
                        self.fields[ annot['T'] ] = annot['V']
                        # Alternative choices in annot['Opt'] )
                    else:
                        raise Exception( "Unknown Widget" )
            else:
                raise Exception( "Unknown Annotation" )
    def _get_text( self, device ):
        text= []
        layout = device.get_result()
        for obj in layout:
            if isinstance( obj, LTTextBoxHorizontal ):
                if obj.get_text().strip():
                    text.append( TextBlock(obj.x0, obj.y1, obj.get_text().strip()) )
        text.sort( key=lambda row: (-row.y, row.x) )
        return text
    def is_recognized( self ):
        """Check for Copyright as well as Revision information on each page."""
        bottom_page_1 = self.text[1][-3:]
        bottom_page_2 = self.text[2][-3:]
        pg1_rev= "Rev 2011.01.17" == bottom_page_1[2].text
        pg2_rev= "Rev 2011.01.17" == bottom_page_2[0].text
        return pg1_rev and pg2_rev