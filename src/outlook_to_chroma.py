import win32com.client
import os
import glob
import olefile

from urllib.parse import unquote
import email.charset
from email.message import EmailMessage

import zipfile
import docx
from pptx import Presentation
import pandas as pd
import chromadb
from chromadb.utils import embedding_functions

from langchain.schema.document import Document
from langchain.document_loaders import TextLoader, PyPDFLoader, CSVLoader
from langchain.embeddings.openai import OpenAIEmbeddings
from langchain.text_splitter import CharacterTextSplitter
from langchain.vectorstores.chroma import Chroma

import sys
sys.path.append('C:/Users/dwchoi0610/dongwook/langchain/outlook/outlook_pipeline/src')

from utils.utils import hwp_to_txt

os.environ['openai_api_key'] = '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'


attachment_folder = 'C:/Temp/outlook/attachments'
eml_folder = 'C:/Temp/outlook/eml'

def delete_files(folder:str, extension: str):
    for f in glob.glob(f"{folder}/*.{extension}"):
        os.remove(f)

def fetch_outlook():
    # fetch application from outlook application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # 6 means INBOX folder
    inbox = outlook.GetDefaultFolder(6)
    return inbox.Items

def make_folders():
    try:
        os.makedirs('C:/Temp/outlook/attachments')
    except Exception:
        print("attachment folder already created")
    
    try:
        os.makedirs('C:/Temp/outlook/eml')
    except Exception:
        print("eml folder already created")


def attachment_parsing(docs, attachment_path, filename, metadata, text_splitter):

    if filename.endswith("pdf"):
        loader = PyPDFLoader(attachment_path)
        loaded = loader.load()
        for i, doc in enumerate(loaded):
            loaded[i].metadata.update(metadata)

        if isinstance(loaded[0], list):
            print(filename)
        # print(type(loaded[0]))
        docs.extend(loaded)

    elif filename.endswith(("docx", "doc")):
        doc = docx.Document(attachment_path)
        text = ' '.join([paragraph.text for paragraph in doc.paragraphs])
        
        if isinstance([Document(page_content=x, metadata = metadata) for x in text_splitter.split_text(text)][0], list):
            print(filename)
        # print(type([Document(page_content=x, metadata = metadata) for x in text_splitter.split_text(text)][0]))
        docs.extend([Document(page_content=x, metadata = metadata) for x in text_splitter.split_text(text)])

    elif filename.endswith(("xlsx", "xls")):
        df = pd.read_excel(attachment_path)
        filename = filename.split('.')[0]+".txt"
        txt_path = attachment_path.split('.')[0] + ".txt"
        df.to_csv(txt_path, sep='\t', index=False, encoding='cp949')
        attachment_parsing(docs, txt_path, filename, metadata, text_splitter)

    elif filename.endswith("csv"):
        loader = CSVLoader(attachment_path)
        loaded = loader.load()
        for i, doc in enumerate(loaded):
            loaded[i].metadata.update(metadata)

        if isinstance(loaded[0], list):
            print(filename)
        # print(type(loaded[0]))
        docs.extend(loaded)

    elif filename.endswith("hwp"):
        f = olefile.OleFileIO(attachment_path)  
        encoded_text = f.openstream('PrvText').read() 
        decoded_text = encoded_text.decode('utf-16')

        if isinstance([Document(page_content=x, metadata = metadata) for x in text_splitter.split_text(decoded_text)][0], list):
            print(filename)
        # print(type([Document(page_content=x, metadata=metadata) for x in text_splitter.split_text(decoded_text)][0]))
        docs.extend([Document(page_content=x, metadata=metadata) for x in text_splitter.split_text(decoded_text)])

    elif filename.endswith(("pptx", "ppt")):
        prs = Presentation(attachment_path)
        text = ''
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text'):
                    text += shape.text + '\n'
        if isinstance([Document(page_content=x, metadata=metadata) for x in text_splitter.split_text(text)][0], list):
            print(filename)
        # print(type([Document(page_content=x, metadata=metadata) for x in text_splitter.split_text(text)][0]))
        docs.extend([Document(page_content=x, metadata=metadata) for x in text_splitter.split_text(text)])

    elif filename.endswith("zip"):
        with zipfile.ZipFile(attachment_path, 'r') as zip_ref:
            file_names = zip_ref.namelist()
            zip_ref.extractall(attachment_folder)
            for file in file_names:
                attachment_path = os.path.join(attachment_folder, file)
                attachment_parsing(docs, attachment_path, file, metadata, text_splitter)

    elif filename.endswith("txt"):
        loader = TextLoader(attachment_path)
        loaded = loader.load()
        for i, doc in enumerate(loaded):
            loaded[i].metadata.update(metadata)
        
        if isinstance(loaded[0], list):
            print(filename)
        # print(type(loaded[0]))
        docs.extend(loaded)

    else:
        pass


def export_data(items):
    make_folders()
    docs = list()

    for i, message in enumerate(items):
        # print("entry id", message.EntryID)
        # print("conv id", message.ConversationID)
        # print("get ids of names", message.GetIDsOfNames)
        
        body = unquote(message.Body)
        text_splitter = CharacterTextSplitter(chunk_size=1000, chunk_overlap=100)

        # metadata
        
        recipients_info = str()
        for idx in range(message.Recipients.Count): 
            recipient = message.Recipients.Item(idx+1)
            recipients_info += f"{recipient.Name} <{recipient.Address}>, "

        metadata = {
            "Subject": str(message.Subject),
            "From": f"{message.SenderName} <{message.SenderEmailAddress}>",
            "To": recipients_info,
            "Date": str(message.ReceivedTime.strftime("%a, %d %b %Y %H:%M:%S %z")),
            "Attchment": False
        }

        # docs for body
        docs.extend([Document(page_content=x, metadata = metadata) for x in text_splitter.split_text(body)])

        # docs for attachments
        for attachment in message.Attachments:
            attachment_path = os.path.join(attachment_folder, attachment.FileName)
            attachment.SaveAsFile(attachment_path)
            filename = attachment.FileName
            metadata['Attchment'] = filename
            attachment_parsing(docs, attachment_path, filename, metadata, text_splitter)

    return docs




if __name__ == "__main__":
    items = fetch_outlook()
    docs = export_data(items)

    embeddings = OpenAIEmbeddings()

    vectordb = Chroma.from_documents(documents=docs, 
                               embedding=embeddings,
                               collection_name="outlooks",
                               persist_directory="chroma_folder")
    
    vectordb.persist()
    