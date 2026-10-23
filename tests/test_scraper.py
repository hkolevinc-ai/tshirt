import unittest
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from scraper import (clean_description, original_image, composition,
                     size_measurements, extract_page, variants, classify,
                     write_template, QN)


class DataRules(unittest.TestCase):
    def test_logistics_removed_without_losing_composition(self):
        value = clean_description('<p>Тениска с щампа.</p><p>Състав: 100% памук.</p><p>Доставка със Спиди.</p><p>Връщане до 14 дни.</p><p>Цената е 12 евро.</p>')
        self.assertIn('100% памук', value)
        self.assertNotIn('Спиди',value)
        self.assertNotIn('14 дни',value)
        self.assertNotIn('12 евро',value)

    def test_original_urls_not_thumbnails(self):
        expected='https://tshirt24.bg/image/catalog/catalog/products/a.png'
        self.assertEqual(original_image('https://tshirt24.bg/catalog/catalog/products/a.png'),expected)
        self.assertEqual(original_image('https://tshirt24.bg/image/cache/catalog/catalog/products/a-694x694w.png'),expected)
        self.assertEqual(original_image('catalog/catalog/products/a.png'),expected)

    def test_composition_is_not_invented(self):
        self.assertEqual(composition('<p>80% памук, 20% полиестер</p>'),{'Cotton':80.0,'Polyester':20.0})
        self.assertEqual(composition('80% памук'),{})

    def test_flat_width_doubles_for_chest_only(self):
        source='<table><tr><td>МЪЖКА БЛУЗА</td><td>S</td><td>M</td></tr><tr><td>Широчина</td><td>51</td><td>56</td></tr><tr><td>Височина</td><td>67</td><td>70</td></tr></table>'
        self.assertEqual(size_measurements(source,'M'),{'Chest':112.0,'Length':70.0})

    def test_live_product_identity_must_match(self):
        with self.assertRaises(ValueError):
            extract_page('<h1>404</h1>','1')

    def test_two_colors_three_sizes_make_six_unique_combinations(self):
        groups=[{'name':'Цвят','choices':[{'id':'1','label':'Бял'},{'id':'2','label':'Черен'}],'required':True,'custom':False},
                {'name':'Мъжки размер','choices':[{'id':str(i),'label':s} for i,s in enumerate(['S','M','L'])],'required':True,'custom':False}]
        result=list(variants({'export_options':[{}]},{'groups':groups},100))
        self.assertEqual(len(result),6)
        with self.assertRaises(ValueError):
            list(variants({'export_options':[{}]},{'groups':groups},5))

    def test_set_is_not_misclassified_as_womens_two_piece_top(self):
        self.assertIsNone(classify({'name(bg)':'Комплект тениски за мама, татко и дете'})[0])
        self.assertEqual(classify({'name(bg)':'Мъжка бяла тениска'})[0],'30469')
        self.assertEqual(classify({'name(bg)':'Дамски суитчър с качулка'})[0],'29065')

    def test_output_serializer_preserves_reserved_rows_and_other_parts(self):
        # A minimal OOXML ZIP fixture, rather than a user workbook.
        with tempfile.TemporaryDirectory() as td:
            src, dst = Path(td)/'input.zip', Path(td)/'output.zip'
            rows=''.join(f'<row r="{i}"><c r="A{i}" t="inlineStr"><is><t>RESERVED{i}</t></is></c></row>' for i in range(1,5))
            with zipfile.ZipFile(src,'w') as z:
                z.writestr('xl/workbook.xml','<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Template" r:id="r1"/></sheets></workbook>')
                z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>')
                z.writestr('xl/worksheets/sheet1.xml','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A1"/><sheetData>'+rows+'<row r="5"><c r="A5" s="7"/></row></sheetData><extLst>KEEP</extLst></worksheet>')
                z.writestr('xl/worksheets/hidden.xml',b'KEEP EXACT BYTES')
            write_template(src,dst,[{'cells':{0:'Крем & бяло',1200:14.99}}])
            with zipfile.ZipFile(dst) as z:
                self.assertEqual(z.read('xl/worksheets/hidden.xml'),b'KEEP EXACT BYTES')
                xml=z.read('xl/worksheets/sheet1.xml').decode()
                self.assertIn(rows,xml)
                self.assertIn('<extLst>KEEP</extLst>',xml)
                root=ET.fromstring(xml)
                self.assertEqual(len(root.find(QN('sheetData'))),5)
                cells=list(root.find(QN('sheetData'))[-1])
                self.assertEqual(cells[0].attrib['s'],'7')
                self.assertEqual(cells[0].find(QN('is')).find(QN('t')).text,'Крем & бяло')
                self.assertEqual(float(cells[1].find(QN('v')).text),14.99)


if __name__=='__main__':
    unittest.main()
