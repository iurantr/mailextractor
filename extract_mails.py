# -*- coding: utf-8 -*-
#!/usr/bin/env python
#
# Extracts email addresses from one or more .pdf, .doc, .docx files.
#
# Can take a folder as input to proccess all the files in it (not in the sub-folders)
#
# Without arguments, processes all the subfolders in the 
# current folder (1-subfolder depth) + current_folder itself.
#
# Output is saved in the foldername_emails.txt
# The z_log_emails_... .txt file with statistics and errors is created when crocessing the folder.
#
# 2019 Iurii Antropov <yuantrop@yahoo.com>
#
# Built upon the scripts of:
# (c) 2013  Dennis Ideler <ideler.dennis@gmail.com>
# https://gist.github.com/dideler/5219706
#
# And of vinovator:
# https://gist.github.com/vinovator/c78c2cb63d62fdd9fb67


### includes===============================================================
from optparse import OptionParser
import os.path
import re

import subprocess # to launch some console utilities and retrieve at output
import ntpath # To get the name from the path no matter the operating system

import docx2txt #For extraction from docx

#For extration from pdf
''' Important classes to remember
PDFParser - fetches data from pdf file
PDFDocument - stores data parsed by PDFParser
PDFPageInterpreter - processes page contents from PDFDocument
PDFDevice - translates processed information from PDFPageInterpreter to whatever you need
PDFResourceManager - Stores shared resources such as fonts or images used by both PDFPageInterpreter and PDFDevice
LAParams - A layout analyzer returns a LTPage object for each page in the PDF document
PDFPageAggregator - Extract the decive to page aggregator to get LT object elements
'''
    
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
# From PDFInterpreter import both PDFResourceManager and PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
#from pdfminer.pdfdevice import PDFDevice
# Import this to raise exception whenever text extraction from PDF is not allowed
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
from pdfminer.converter import PDFPageAggregator


#Configure the PATH to the antiword package
import sys
antiword_path="C:\\"
if getattr( sys, 'frozen', False ) :
        # if running in a bundle  (as .exe)
        antiword_path=sys._MEIPASS # search for the antiword in the tmp folder of the extracted bundle
        #print("antiword_path = " + antiword_path)
        os.environ["PATH"] += os.pathsep + os.path.join(antiword_path, "antiword")
        os.environ["HOME"] = antiword_path
        #print(os.environ["PATH"])     
else :
        # running live (as .py)
        # search for antiword subfolder in scripts current folder
        script_parent_dir = os.path.dirname(__file__)
        antiword_path = os.path.join( script_parent_dir ,"antiword")
        if os.path.isdir(antiword_path):
            os.environ["PATH"] += os.pathsep + antiword_path
            os.environ["HOME"] = script_parent_dir
        else:
            #Assuming that antiword is properly installed 
            pass
        #print("antiword_path = " + antiword_path)
        
### End includes===============================================================

# Regex to search for emails in the plain text
regex = re.compile(("([a-z0-9][a-z0-9!#$%&'*+\/=?^_`{|}~-]*"
                    "(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*"
                    "@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.))+"
                    "[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)"))

##regex to match the line-spaced emails: t e s t . t e s t 2 @ g m a i l . c o m
#(it's not used anyware in the code yet)
regex2 = re.compile("\s((?:[a-z0-9!#$%&'*+\/=?^_`{|}~-] )+(?:(?:\. )(?:[a-z0-9!#$%&'*+\/=?^_`{|}~-] )+)*@ (?:[a-z0-9] (?:(?:[a-z0-9-] )*[a-z0-9] )?(?:\. ))+[a-z0-9] (?:(?:[a-z0-9-] )*[a-z0-9])?)\s")


# regex which match the word "dot" in place of "." and word "at" instead of "@" like:
# emailName at gmail dot com
#regex = re.compile(("([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`"
#                    "{|}~-]+)*(@|\sat\s)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|"
#                    "\sdot\s))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)"))

                        
def file_to_str(filename):
    """Returns the contents of filename as a string."""
    with open(filename, "r") as f:
        return f.read().lower() # Case is lowered to prevent regex mismatches.

def doc_to_str(filename):
    "Return a contents of the .doc file as string. Uses antiword."
    #antiword should be installed and configured in the system
    #antiword recognizes it the .doc file is zipped and extracts it.
    #yet, there are some issues with bad recognition of headers/footnotes in the documents
    #and, maybe, some other special elements
    MyOut = subprocess.Popen(["antiword", filename], 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT)
    stdout,stderr = MyOut.communicate()
    #print(stdout.decode("utf-8", errors="ignore"))
    return stdout.decode("utf-8", errors="ignore")

def doc_as_txt_to_str(filename):
    """Returns the contents of the .doc file reading it as a plain text file.
Ignores encoding errors"""
    #- It doesn't work on all the .doc files, some of them are zipped.
    #- Works well on headers and footnotes
    #- Known issue: some non-text gibberish might be later matched by the regex,
    # like: at7@x.c
    with open(filename, "rb") as f:
        return f.read().lower().decode("utf-8", errors='ignore') # Case is lowered to prevent regex mismatches.

# Textract - powerfull library to extract text from many document formats.
# Yet, somestimes gives an encoding error on processing the .doc files
#import textract      
#def doc_to_str(filename):
#    """Returns the contents of Microsoft Word .doc with name <filename> as a string."""
#    return textract.process(filename).decode("utf-8").lower()   

#======================================================================================= 
def pdf_to_str(pdf_filepath):
    """Returns the contents of pdf as a string."""
    
    # Code is taken and modified from:
    # https://gist.github.com/vinovator/c78c2cb63d62fdd9fb67
    
    # pdfTextMiner.py
    # Python 2.7.6
    # For Python 3.x use pdfminer3k module
    # This link has useful information on components of the program
    # https://euske.github.io/pdfminer/programming.html
    # http://denis.papathanasiou.org/posts/2010.08.04.post.html
    
    ''' This is what we are trying to do:
    1) Transfer information from PDF file to PDF document object. This is done using parser
    2) Open the PDF file
    3) Parse the file using PDFParser object
    4) Assign the parsed content to PDFDocument object
    5) Now the information in this PDFDocumet object has to be processed. For this we need
       PDFPageInterpreter, PDFDevice and PDFResourceManager
     6) Finally process the file page by page 
    '''
    
#    my_file = os.path.join("./" + pdf_filepath)
    
    password = ""
    extracted_text = ""
    
    # Open and read the pdf file in binary mode
    fp = open(pdf_filepath, "rb")
    
    # Create parser object to parse the pdf content
    parser = PDFParser(fp)
    
    # Store the parsed content in PDFDocument object
    document = PDFDocument(parser, password)
    
    # Check if document is extractable, if not abort
    if not document.is_extractable:
    	raise PDFTextExtractionNotAllowed
    	
    # Create PDFResourceManager object that stores shared resources such as fonts or images
    rsrcmgr = PDFResourceManager()
    
    # set parameters for analysis
    laparams = LAParams()
    
    # Create a PDFDevice object which translates interpreted information into desired format
    # Device needs to be connected to resource manager to store shared resources
    # device = PDFDevice(rsrcmgr)
    # Extract the decive to page aggregator to get LT object elements
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    
    # Create interpreter object to process page content from PDFDocument
    # Interpreter needs to be connected to resource manager for shared resources and device 
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Ok now that we have everything to process a pdf document, lets process it page by page
    
    for page in PDFPage.create_pages(document):
    	# As the interpreter processes the page stored in PDFDocument object
    	interpreter.process_page(page)
    	# The device renders the layout from interpreter
    	layout = device.get_result()
    	# Out of the many LT objects within layout, we are interested in LTTextBox and LTTextLine
    	for lt_obj in layout:
    		if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
    			extracted_text += lt_obj.get_text()
    			
    #close the pdf file
    fp.close()
    
    # print (extracted_text.encode("utf-8"))
    return extracted_text

### End of pdf_to_str() =======================================================

def get_emails(s):
    """Returns an list of matched emails found in string s."""
    # Removing lines that start with '//' because the regular expression
    # mistakenly matches patterns like 'http://foo@bar.com' as '//foo@bar.com'.
    return [email[0] for email in re.findall(regex, s) if not email[0].startswith('//')]
    
def extract_emails_from_file(filepath):
    "Read a file according to it's extension and return extracted emails "
    print("Processing file: " + filepath)
    filename, file_extension = os.path.splitext(filepath)        
    emails=[]
# Let's drop txt files support for now
#    if file_extension=='.txt':
#        for email in get_emails(file_to_str(filepath)):
#            emails.append(email)
#            print(email)
    if file_extension=='.pdf':
        for email in get_emails(pdf_to_str(filepath)):
            emails.append(email)
            print(email)
    elif file_extension=='.docx':
        text = docx2txt.process(filepath)
        for email in get_emails(text):
            emails.append(email)
            print(email)
    elif file_extension=='.doc':
        text = doc_to_str(filepath) # try antiword tool
        emails=get_emails(text)
        
        if len(emails)==0 and "is not a Word Document." in text:
            #antiword doesn't recognize more recent .doc.
            #try docx2txt in case the .doc is zipped .xml file
            text = docx2txt.process(filepath)
            emails=get_emails(text)
        if len(emails)==0:
            #if nothing have helped, try to open .doc file as a plain text
            #this method is prone to errors: some random noise may look like an email, like:
            # 4j30@hsx.c
            text=doc_as_txt_to_str(filepath)
            emails=get_emails(text)
            
        for email in emails:
            print(email)
    else:
        #print("L'extension de fichier est inconnue!!!")
        pass
    return emails

def path_leaf(path):
    "Return the filename.ext or dirname.ext from the path"
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

def filter_unique_emails(sequence):
    "Clear the duplicates from sequence, but keep an order"
    seen = set()
    return [x for x in sequence if not (x in seen or seen.add(x))]   

###  Main =====================================================================
if __name__ == '__main__':
    #Print Usage if run from console
    parser = OptionParser(usage=("\npython %prog [FILE]...\n"
                                 "python %prog [DIR]...\n"
                                 "python %prog <-- to process the same directory")
                                )
    # No options added yet. Add them here if you ever need them.
    options, args = parser.parse_args()

    #If no input arguments -  run on all the folders in the current directory
    #and on the current directory
    if not args:
        #args=[x in os.listdir('./')
        args=[x for x in os.listdir('./') if os.path.isdir('./'+ x)]
        args.append("./")
   
    #The list of files or directories is expected as an arguments
    for arg in args:
        #For counters and extracted emails
        all_emails=[]
        unique_emails=set()
        files_processed=0
        files_unknown_format=[]
        files_without_emails=[]
        files_with_emails=0
        files_with_only_duplicate_mails=[]
        files_with_errors=[]
        duplicate_emails=0
        
        #If arg is a file
        if os.path.isfile(arg):
            filename, file_extension = os.path.splitext(arg)
            if file_extension not in ['.pdf', '.doc', '.docx']:
                continue
            try:
                all_emails+=extract_emails_from_file(arg)
                unique_emails = filter_unique_emails(all_emails)
            except Exception as e:
                print("Error processing file: {}".format(arg))
                print(e)
                continue
            #print("Finis !!")    
        #If arg is a folder
        elif os.path.isdir(arg):
            # temporary log file with extracted emails. It is deleted on success.
            with open(arg + "_emails_extract_tmp.log","w", encoding="utf-8") as log_file:
                n_files=0
                for file in os.listdir(arg):
                    filepath=arg + "/" + file
                    if os.path.isfile(filepath):
                        n_files+=1
                        filepathname, file_extension = os.path.splitext(filepath)
                        filename = path_leaf(filepath)
                        if file_extension not in ['.pdf', '.doc', '.docx']:
                            files_unknown_format.append(filename)
                            continue                       
                        try:
                            emails=extract_emails_from_file(filepath)
                        except Exception as e:
                            print("Error processing file: {}".format(filepath))
                            print(e)
                            files_with_errors.append(filename)
                        else:
                            files_processed+=1
                            all_emails+=emails

                            if len(emails)==0:
                                files_without_emails.append(filename) 
                            else:
                                new_emails=0
                                for email in emails:
                                    if not email in unique_emails:
                                        unique_emails.add(email)
                                        new_emails+=1
                                if new_emails==0:
                                    files_with_only_duplicate_mails.append(filename)
                                else:
                                    files_with_emails+=1
                            log_file.write("{};{}\n".format(file, emails))
                            if (n_files % 10) ==0:
                                log_file.flush()
                            
        else:
            print('"{}" is not a file or directory.'.format(arg))
            parser.print_usage()
            continue
    
        #Write an output
        outputName=arg+"_emails.txt"
        dirname=os.path.dirname(os.path.abspath(arg))
        logName = dirname + "/z_log_emails_" + path_leaf(arg)+ ".txt"
        if arg=="./":
            outputName="./" + path_leaf(os.path.abspath("./")) + "_curr_folder_emails.txt"
            logName =  dirname + "/z_log_emails_" + path_leaf(os.path.abspath("./"))+ "_curr_folder.txt"
            
        if len(all_emails)!=0:
            with open(outputName, 'w') as f:
#                filtered_emails = unique_emails(all_emails)
                f.write("\n".join(unique_emails) + "\n")
                
            #Print statistics on the screen
            print("Unique emails extracted: {}".format( len(unique_emails) ) )
            if (os.path.isdir(arg)):
                print("Files total: {}".format( files_processed ) )
                print("Files with unknown extension: {}".format( len(files_unknown_format) ) )
                print("Files with new emails: {}".format( files_with_emails ) )
                print("Files with duplicate emails: {}".format( len(files_with_only_duplicate_mails )))
                print("Files without emails: {}".format( len(files_without_emails )) )
                print("Files with errors: {}".format( len(files_with_errors )))
                
                #And write them in the separate <z_log_... .txt> file
                #Also, list all the files with errors, with no emails, and with only duplicates

                dirname=os.path.dirname(os.path.abspath(arg))
                with open(logName ,'w') as log_file:
                    log_file.write("Unique emails extracted: {}\n".format( len(unique_emails) ))
                    log_file.write("Files processed: {}\n".format( files_processed ) )
                    log_file.write("Files with unsupported extension: {}\n".format( len(files_unknown_format) ) )
                    log_file.write("Files with new emails: {}\n".format( files_with_emails ) )
                    log_file.write("Files with duplicate emails: {}\n".format( len(files_with_only_duplicate_mails )))
                    log_file.write("Files without emails: {}\n".format( len(files_without_emails )) )
                    log_file.write("Files with errors: {}\n\n".format( len(files_with_errors )))
                    
                    if len(files_with_errors)>0:
                        log_file.write("FILES WITH ERRORS:\n")
                        log_file.write("\n".join(files_with_errors) + "\n\n")
                        
                    if len(files_unknown_format)>0:
                        log_file.write("FILES WITH UNSUPPORTED EXTENSION:\n")
                        log_file.write("\n".join(files_unknown_format) + "\n\n")
                        
                    if len(files_without_emails)>0:
                        log_file.write("FILES WITHOUT EMAILS:\n")
                        log_file.write("\n".join(files_without_emails) + "\n\n")     
                    
                    if len(files_with_only_duplicate_mails)>0:
                        log_file.write("FILES WITH DUPLICATE EMAILS:\n")
                        log_file.write("\n".join(files_with_only_duplicate_mails) + "\n\n")      
            
        #Delete temp file if everything finished successfully
        if os.path.isfile(arg + "_emails_extract_tmp.log"):
            os.remove(arg + "_emails_extract_tmp.log")
                
    
            