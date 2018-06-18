import os
import re
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfparser import PDFDocument, PDFParser
from regex_statements import RegexStatements


def move_file1(path):
    """Organises files to folder per supplier.

       Uses regular expressions to id the files.
    
       path: root folder for all saved order confirmations.
    """    
    cond_dict = {"Agriservice": RegexStatements.agriservice,
                #  "Alubel": RegexStatements.alubel,
                 "Aris": RegexStatements.aris,
                 "Banner": RegexStatements.banner,
                 "Berardi": RegexStatements.berardi,
                #  "Bertolini": RegexStatements.bertolini,
                #  "BKS": RegexStatements.bks,
                 "Casoli-Losi": RegexStatements.casoli,
                 "CCRE": RegexStatements.ccre,
                 "CFR": RegexStatements.cfr,
                 "Conduxtix-Wampfler": RegexStatements.conductix,
                 "Crimo": RegexStatements.crimo,
                 "Dalmar": RegexStatements.dalmar,
                 "Dram": RegexStatements.dram,
                 "Emec": RegexStatements.emec,
                 "Emka": RegexStatements.emka,
                 "Fimo": RegexStatements.fimo,
                 "Flexpo": RegexStatements.flexpo,
                 "FM": RegexStatements.fm,
                 "Franzini": RegexStatements.franzini,
                 "GMP": RegexStatements.gmp,
                 "Hilti": RegexStatements.hilti,
                 "IFM": RegexStatements.ifm,
                 "Imytech": RegexStatements.imytech,
                #  "Interpump": RegexStatements.interpump,
                 "Itap": RegexStatements.itap,
                 "Leadercom": RegexStatements.leadercom,
                 "Levratti": RegexStatements.levratti,
                 "Liverani": RegexStatements.liverani,
                 "Lombardini": RegexStatements.lombardini,
                 "MB Italia": RegexStatements.mbitalia,
                #  "Motorgarden": RegexStatements.motorgarden,
                #  "Mecc 2000": RegexStatements.mecc,
                 "Noyfar": RegexStatements.noyfar,
                 "OMZ": RegexStatements.omz,
                #  "Oleomeccanica": RegexStatements.oleomeccanica,
                 "PA": RegexStatements.pa,
                 "Paluan": RegexStatements.paluan,
                 "PBM": RegexStatements.pbm,
                 "Pepperl+Fuchs": RegexStatements.pepperl,
                 "PFA": RegexStatements.pfa,
                 "Planet Filters": RegexStatements.planetfilers,
                 "PPE": RegexStatements.ppe,
                 "P Service": RegexStatements.pservice,
                 "SAI Electric": RegexStatements.sai,
                #  "Sauro Rossi": RegexStatements.sauro,
                 "Simonazzi Remo": RegexStatements.simonazzi,
                 "Sintostamp": RegexStatements.sintostamp,
                 "Telcom": RegexStatements.telcom,
                 "Tecomec": RegexStatements.tecomec,
                 "Tellarini": RegexStatements.tellarini,
                 "Transtecno": RegexStatements.transtecno,
                #  "Tubi Gomma Torino": RegexStatements.tgt,
                #  "Thenar": RegexStatements.thenar,
                 "Vetroresina Padana": RegexStatements.vetroresina,
                 "Wicke": RegexStatements.wicke,
                 "Xylem": RegexStatements.xylem,
                 "Zeli": RegexStatements.zeli
                 }

    # Create a list of files to be organised.
    with os.scandir(path) as files:
        for file in files:
            if file.is_file():
                for supplier in cond_dict:  # Get suppliers from dictionary (keys)
                    for cond in cond_dict[supplier]:  # Regex conditions are in lists, so loop through them if more than 1 cond.
                        regex_cond = re.compile(cond)
                        filename = file.name
                        file_path = file.path
                        dst_file = os.path.join(path, supplier, filename)
                        if regex_cond.search(filename):  # Compare regex statement to filename
                            create_folder(path, supplier)
                            rename_file(file_path, dst_file, filename, path, supplier)


def move_file2(path):
    """Organises files to folder per supplier.

       Uses pdfparser to id the files.
    
       path: root folder for all saved order confirmations.
    """    

    # Create a list of files to be organised.
    with os.scandir(path) as files:
        for file in files:
            if file.is_file() and file.name[-3:].lower() == 'pdf':  # Handle only pdf files
                filename = file.name
                file_path = file.path
                supplier, nr_conf, date_conf, nr_ord = fileparser(file_path, path)
                if supplier:  # If supplier variable is available, it means that a pdf was parsed correctly
                    dst_file = os.path.join(path, supplier, filename)
                    create_folder(path, supplier)
                    rename_file(file_path, dst_file, filename, path, supplier, nr_conf, date_conf, nr_ord)
                    

def create_folder(path, supplier):
    if not os.path.exists(os.path.join(path, supplier)):
        os.mkdir(os.path.join(path, supplier))


def rename_file(file_path, dst_file, filename, path, supplier, nr_conf="", date_conf="", nr_ord=""):
    """Renames files

       Moves files to company specific folders and renames the file with order nr etc if needed.

       Args:
       file_path: path to the file to be handled
       dst_file: path to the new locations of the file
       filename: current file name
       path: path to the root folder where order confirmations are stored
       supplier: Owner company of the document
       nr_conf: Order confirmation number
       date_conf: Order confirmation date
       nr_ord: Order number (our company's)
    """
    count = 1
    if nr_conf:
        nr_conf = nr_conf.replace('/', '-')
        date_conf = date_conf.replace('/', '-')
        date_conf = date_conf.replace('.', '-')
        filename = "{}{}".format("_".join([nr_conf, date_conf, nr_ord]), ".pdf")
        ls = dst_file.split('\\')[:-1]
        ls.append(filename)
        dst_file = "\\".join(ls)
    while os.path.exists(dst_file):  # If destination file exists, add a progressive number to filename (001 -> ...)
        head, tail = os.path.split(dst_file)
        ext = tail[-4:]  # ext should be .pdf
        len_ext = len(ext)
        if count > 1:
            tail = tail[:-len_ext-4]
        else:
            tail = tail[:-len_ext]
        dst_file = "{}\\{}-{}{}".format(head, tail, str(count).zfill(3), ext)
        count += 1
    os.rename(file_path, dst_file)
    if count > 1:
        print("The file {} was already in {}{}.\nIt was renamed to {}-{}{}.".format(
            filename, path, supplier, tail, str(count-1).zfill(3), ext))


def fileparser(file, path):
    """Parses info from order confirmations

       Args:
       file: path to the file to be handled
       path: path to the root folder where the pdf files are stored

       Output:
       supplier: Owner company of the document
       nr_conf: Order confirmation number
       date_conf: Order confirmation date
       nr_ord: Order number (our company's)
    """
    print(file)
    temp_file = "tmp"
    with open(file, mode='rb') as f:

        parser = PDFParser(f)
        doc = PDFDocument()
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize('')
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        # I changed the following 2 parameters to get rid of white spaces inside words:
        laparams.char_margin = 1.0
        laparams.word_margin = 1.0
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        extracted_text = ''

        # Process each page contained in the document.
        for page in doc.get_pages():
            interpreter.process_page(page)
            layout = device.get_result()
            for lt_obj in layout:
                # Extract the text
                if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                    extracted_text += str(lt_obj)

        with open(os.path.join(path, temp_file), "wb") as txt_file:
            txt_file.write(extracted_text.encode("utf-8"))      

    
    with open(os.path.join(path, temp_file), "r") as txt_file:
        supplier = ""
        nr_conf = ""
        date_conf = ""
        nr_ord = ""
        # Cleaning up the the parsed data
        lines = txt_file.read()
        lines = lines.split('><')
        for i in range(len(lines)):
            lines[i] = lines[i].strip('<')
            lines[i] = lines[i].split(" ")
            lines[i] = "".join(lines[i])
            lines[i] = lines[i].split("\\n")
        # Conditions for recognize to which company the file belongs to
        # and positions for other data
        if lines[0][0] == r"LTTextBoxHorizontal(0)19.600,663.672,475.792,736.296'SPETT.LE":
            supplier = "Mecc2000" 
            nr_conf = lines[3][0][62:68] 
            date_conf = lines[3][1][10:20]
            nr_ord = lines[7][0][64:83]
        if lines[0][0] == r"LTTextBoxHorizontal(0)226.004,760.963,256.439,777.959'SRL":
            supplier = "BKS" 
            nr_conf = lines[10][0][53:59] 
            date_conf = lines[16][0][55:65]
        if lines[0][0] == r"LTTextBoxHorizontal(0)56.640,766.847,338.754,789.243'MOTORGARDENs.r.l.":
            supplier = "Motor Garden"
            nr_conf = lines[13][0][55:58]
            date_conf = lines[12][0][55:65]
            nr_ord = lines[9][0][53:68]
        if lines[0][0][-12:] == r"Numeroordine":
            supplier = "Sauro Rossi"
            nr_conf = lines[4][0][52:]
            date_conf = lines[5][0][54:64]
            nr_ord = lines[35][0][-3:]
        if lines[0][0] == r"LTTextBoxHorizontal(0)314.150,735.539,421.206,796.939'Spett.le":
            supplier = "SAI Electric"
            nr_conf = lines[3][0][-6:]
            date_conf = lines[5][0][54:]

            
    os.remove(os.path.join(path, temp_file))
    return supplier, nr_conf, date_conf, nr_ord

    
def main():
    path = "C:\\Users\\Utente-006\\Documents\\Programming\\PdfOrganiser\\PDF\\"
    move_file1(path)
    move_file2(path)


if __name__ == "__main__":
    main()
