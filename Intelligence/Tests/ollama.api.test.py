from langchain_ollama import OllamaLLM
from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser

# Initialize the Ollama LLM
llm = OllamaLLM(model="llama3.2:3b")  # Replace "llama2" with your model name if needed

# Create a prompt template
template = """Question: {question}

Answer: Let's think step by step."""
prompt = PromptTemplate.from_template(template)

# Create a chain for question-answering
chain = prompt | llm | StrOutputParser()

# Run the chain
response = chain.invoke({"question": "What is LangChain?"})
print(response)