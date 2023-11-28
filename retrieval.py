import openai
import pandas as pd
import os

import chromadb

from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores.chroma import Chroma

from langchain.prompts import PromptTemplate
from langchain.agents.openai_functions_agent.base import OpenAIFunctionsAgent
from langchain.prompts import MessagesPlaceholder
from langchain.schema.messages import SystemMessage
from langchain.agents import AgentExecutor

from langchain.agents.agent_toolkits import create_retriever_tool
from langchain.agents.agent_toolkits import create_conversational_retrieval_agent
from langchain.chat_models import ChatOpenAI

from langchain.text_splitter import RecursiveCharacterTextSplitter

# Ignore unclosed SSL socket warnings - optional in case you get these errors
import warnings

warnings.filterwarnings(action="ignore", message="unclosed", category=ResourceWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning) 

################ environment settings (for query embedding) ################
############################################################################

EMBEDDING_MODEL = "text-embedding-ada-002"
os.environ['openai_api_key'] = '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'
api_key = os.environ.get('OPENAI_API_KEY')

if os.getenv("OPENAI_API_KEY") is not None:
    openai.api_key = os.getenv("OPENAI_API_KEY")
    print ("OPENAI_API_KEY is ready")
else:
    print ("OPENAI_API_KEY environment variable not found")

################ Chroma DB ################
###########################################

def chain_load_db():
    # load chroma db
    # chroma_client, collections = open_chroma()
    # chroma_client = open_chroma()
    chroma_client = chromadb.PersistentClient(path="./src/chroma_folder")
    
    # wikipedia_content_collection, wikipedia_title_chroma = collections
    
    # embedding function
    embedding_function = OpenAIEmbeddings(
                            api_key=os.environ.get('OPENAI_API_KEY'),
                            model="text-embedding-ada-002",
                        )
    
    #### pass over chroma to langchain ####
    chroma = Chroma(
        persist_directory="./chroma",
        client=chroma_client,
        collection_name="outlooks",
        embedding_function=embedding_function
    )

    # make retriever
    ## search_kwargs={"k": 5} -> top k 설정
    retriever = chroma.as_retriever(search_kwargs={"k": 5})

    ## make retriever setting
    tool = create_retriever_tool(
        retriever,
        "Search_Outlook_emails",
        "Searches and returns documents regarding email's contents and attachments.",
    )
    tools = [tool]

    # make llm
    llm = ChatOpenAI(model_name="gpt-4-1106-preview", temperature=0)

    # make prompt
    ## SystemMessage 
    ### A Message for priming AI behavior, usually passed in as the first of a sequence
    ### of input messages.

    system_message = SystemMessage(
        content=(
            "너는 레퍼런스 문서들로부터 정보를 얻어 나에게 대답을 해주는 챗봇이야"
            "레퍼런스의 출처와 그 근거를 들어 나에게 답변을 해주면 돼"
        )
    )

    prompt = OpenAIFunctionsAgent.create_prompt(
        system_message=system_message,
    )
    
    # set agent
    agent = OpenAIFunctionsAgent(llm=llm, tools=tools, prompt=prompt)
    agent_executor = AgentExecutor(
        agent=agent,
        tools=tools,
        verbose=True,
        return_intermediate_steps=True,
    )

    return chroma, chroma_client, agent_executor

if __name__ == "__main__":
    chain_load_db()