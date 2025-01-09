from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Dict
from langchain_ollama import OllamaLLM
from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import JsonOutputParser
from langchain_core.output_parsers import StrOutputParser 
from icecream import ic

from layout import LayoutRag
from layout import LayoutPreview
app = FastAPI()
app.include_router(LayoutRag.router)
app.include_router(LayoutPreview.router)


# Ollama API endpoint

class Line(BaseModel):
    text: str


class ClassifiedText(BaseModel):
    classified_lines: Dict[str, List[str] | str]

class LayoutField(BaseModel):
    name : str
    value : int 

class ClassificationRequest(BaseModel):
    lines: List[str]
    classes: List[LayoutField]

@app.post("/classify/", response_model=ClassifiedText)
async def classify_text(request: ClassificationRequest):
    """
    Classifies lines of text into different fields (classes) using Ollama and LangChain.
    """
    llm = OllamaLLM(model="llama3.2:3b")  # Replace "llama2" with your model name if needed

    # Dynamically create the prompt template based on provided classes
    formatted_lines = "\n".join([f"{i+1}. {line}" for i, line in enumerate(request.lines)])
    format = "\n".join([f"{clx.name} : Max {clx.value} Lines" for clx in request.classes ])
    outputFormat = "\n".join([f"\"{clx.name}\" : [\"Item 1\", \"Item 2\", \"Item 3\",]," for clx in request.classes ])

    prompt_template = (
        f"""Using the following lines 
        {{lines}}

        Create a json object that have following fields and the respective amount of data as specified below 

        {{format}}
        
        Also shorten the lines so that they look cool and impactful in a presentation

        Output the data as a json object

        """
    )

    # print(prompt_template)
    # print(outputFormat)
    # print(formatted_lines)
    # print(format)

    prompt = PromptTemplate(
        template=prompt_template,
        input_variables=["lines", "format"]
    )

    chain = prompt | llm | JsonOutputParser()


    # Prepare the lines for the prompt
    try:
        # Run the chain with the formatted lines and class options
        response = chain.invoke({
            "lines": formatted_lines,
            "format": format 
        })

        # Extract the JSON from the response
        classified_data=response
        print(response)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing request: {e}")

    return ClassifiedText(classified_lines=classified_data)

def extract_json_from_ollama_output(output: str) -> Dict:
    """
    Extracts the JSON object from the Ollama output.
    """
    import re
    import json

    # Find the JSON object using regex (adjust the pattern if needed)
    match = re.search(r"```json\n(.*?)\n```", output, re.DOTALL)
    if match:
        json_str = match.group(1)
        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in Ollama output: {e}")
    else:
        raise ValueError("JSON object not found in Ollama output")