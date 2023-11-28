import streamlit as st
from retrieval import chain_load_db
import argparse

st.title("IR search chat bot")
st.caption("streamlit chatbot with IR data in chroma DB")


########## load rag model ##########

def load_client():
    # parser = argparse.ArgumentParser()
    # parser.add_argument('--service', type=str, help='sm, ir')
    # args = parser.parse_args()

    # collection, client, agent_executor 불러오기
    collection, client, agent_executor = chain_load_db()

    st.session_state['client'] = client
    st.session_state['collection'] = collection
    st.session_state['agent_executor'] = agent_executor

if "messages" not in st.session_state:
    load_client()
    st.session_state["messages"] = [{"role": "assistant", "content": "How can I help you?"}]

########## set retriever settings for streamlit ##########

# temperature (x)
# prompting 어떻게 해놨는지? (x)
# => 전부 없고 @로 해야 함
# a = st.sidebar.radio('Choose:',[1,2])


########## chatbot ##########

for msg in st.session_state.messages:
    st.chat_message(msg["role"]).write(msg["content"])

if prompt := st.chat_input():
    st.session_state.messages.append({"role": "user", "content": prompt})
    st.chat_message("user").write(prompt)
    query = st.session_state.messages[-1]['content']
    # get response with query
    result = st.session_state['agent_executor']({"input": query})

    msg = result['output']
    st.session_state.messages.append({"role": "assistant", "content": msg})
    st.chat_message("assistant").write(msg)