import os
from dotenv import load_dotenv

load_dotenv()

# Function to get configuration from Streamlit secrets or environment
def get_config(key, default=None):
    # Check if we're running in Streamlit environment
    is_streamlit = False
    try:
        import streamlit as st
        if hasattr(st, 'secrets'):
            is_streamlit = True
            if key in st.secrets:
                value = st.secrets[key]
                # Skip placeholder values
                if value and not str(value).startswith("your_") and "your_" not in str(value).lower():
                    return value
    except Exception:
        pass
    
    # For local development or fallback, try environment variables
    env_value = os.getenv(key)
    if env_value and env_value != default and not env_value.startswith("your_"):
        return env_value
    
    # If running locally and no env var, try to read from secrets.toml directly
    if not is_streamlit:
        try:
            import toml
            secrets_file = os.path.join(os.path.dirname(__file__), '..', '.streamlit', 'secrets.toml')
            if os.path.exists(secrets_file):
                with open(secrets_file, 'r') as f:
                    secrets = toml.load(f)
                    if key in secrets:
                        value = secrets[key]
                        if value and not str(value).startswith("your_"):
                            return value
        except Exception:
            pass
    
    return default

model_type = get_config('MODEL_TYPE', 'openai')

if model_type == 'ollama':
    from ollama import Client
    
    client = Client(host=get_config('OLLAMA_HOST', 'http://localhost:11434'))
    
    response = client.chat(
        model=get_config('OLLAMA_MODEL', 'llama2'),
        messages=[
            {"role": "user", "content": "write a haiku about ai"}
        ]
    )
    
    print("Ollama",response['message']['content'])
    
elif model_type == 'openai':
    from openai import OpenAI
    
    api_key = get_config('OPENAI_API_KEY')
    print(f"Using API key: {api_key[:10]}***{api_key[-4:] if api_key else 'None'}")
    
    client = OpenAI(
        api_key=api_key
    )
    
    completion = client.chat.completions.create(
        model=get_config('OPENAI_MODEL', 'gpt-4o-mini'),
        store=True,
        messages=[
            {"role": "user", "content": "write a haiku about ai"}
        ]
    )
    
    print(f"OpenAI {get_config('OPENAI_MODEL', 'gpt-4o-mini')}",completion.choices[0].message.content)
    
else:
    raise ValueError(f"Unsupported model_type: {model_type}. Use 'openai' or 'ollama'.")
