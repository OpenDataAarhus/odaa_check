#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Copyright (C) 2007 Søren Roug, European Environment Agency
#
# This is free software.  You may redistribute it under the terms
# of the Apache license and the GNU General Public License Version
# 2 or at your option any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public
# License along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

import unittest
import os
from odf.opendocument import OpenDocumentText, load
from odf import text
from odf.namespaces import TEXTNS

class TestText(unittest.TestCase):
    
    def test_softpagebreak(self):
        """ Create a soft page break """
        textdoc = OpenDocumentText()
        spb = text.SoftPageBreak()
        textdoc.text.addElement(spb)
        self.assertEquals(1, 1)

    def test_1stpara(self):
        """ Grab 1st paragraph and convert to string value """
        poem_odt = os.path.join(
            os.path.dirname(__file__), "examples", "serious_poem.odt")
        d = load(poem_odt)
        shouldbe = u"The boy stood on the burning deck,Whence allbuthim had fled.The flames that litthe battle'swreck,Shone o'er him, round the dead. "
        self.assertEquals(shouldbe, unicode(d.body))
        self.assertEquals(shouldbe, str(d.body))


if __name__ == '__main__':
    unittest.main()
