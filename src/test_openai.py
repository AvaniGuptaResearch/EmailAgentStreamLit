import os
from dotenv import load_dotenv

load_dotenv()

model_type = os.getenv('MODEL_TYPE', 'openai')

if model_type == 'ollama':
    from ollama import Client
    
    client = Client(host=os.getenv('OLLAMA_HOST', 'http://localhost:11434'))
    
    response = client.chat(
        model=os.getenv('OLLAMA_MODEL', 'llama2'),
        messages=[
            {"role": "user", "content": "write a haiku about ai"}
        ]
    )
    
    print("Ollama",response['message']['content'])
    
elif model_type == 'openai':
    from openai import OpenAI
    
    client = OpenAI(
        api_key=os.getenv('OPENAI_API_KEY')
    )
    
    completion = client.chat.completions.create(
        model=os.getenv('OPENAI_MODEL', 'gpt-4o-mini'),
        store=True,
        messages=[
            {"role": "user", "content": "write a haiku about ai"}
        ]
    )
    
    print("OpenAI",completion.choices[0].message.content)
    
else:
    raise ValueError(f"Unsupported model_type: {model_type}. Use 'openai' or 'ollama'.")
