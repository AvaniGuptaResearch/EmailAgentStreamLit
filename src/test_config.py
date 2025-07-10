#!/usr/bin/env python3
"""Test configuration loading for debugging"""

import os
from dotenv import load_dotenv

load_dotenv()

# Function to get configuration from Streamlit secrets or environment
def get_config(key, default=None):
    print(f"Looking for config key: {key}")
    
    # Check if we're running in Streamlit environment
    is_streamlit = False
    try:
        import streamlit as st
        if hasattr(st, 'secrets'):
            is_streamlit = True
            print("Running in Streamlit environment")
            if key in st.secrets:
                value = st.secrets[key]
                # Skip placeholder values
                if value and not str(value).startswith("your_") and "your_" not in str(value).lower():
                    print(f"Found in Streamlit secrets: {key} = {str(value)[:10]}***")
                    return value
                else:
                    print(f"Found placeholder in Streamlit secrets: {key}")
        else:
            print("Streamlit available but no secrets")
    except Exception as e:
        print(f"Streamlit not available: {e}")
    
    # For local development or fallback, try environment variables
    env_value = os.getenv(key)
    if env_value and env_value != default and not env_value.startswith("your_"):
        print(f"Found in environment: {key} = {env_value[:10]}***")
        return env_value
    
    # If running locally and no env var, try to read from secrets.toml directly
    if not is_streamlit:
        try:
            import toml
            secrets_file = os.path.join(os.path.dirname(__file__), '..', '.streamlit', 'secrets.toml')
            print(f"Looking for secrets file: {secrets_file}")
            if os.path.exists(secrets_file):
                print("Found secrets.toml file, reading...")
                with open(secrets_file, 'r') as f:
                    secrets = toml.load(f)
                    if key in secrets:
                        value = secrets[key]
                        if value and not str(value).startswith("your_"):
                            print(f"Found in secrets.toml: {key} = {str(value)[:10]}***")
                            return value
                        else:
                            print(f"Found placeholder in secrets.toml: {key}")
                    else:
                        print(f"Key {key} not found in secrets.toml")
            else:
                print("secrets.toml file not found")
        except Exception as e:
            print(f"Error reading secrets.toml: {e}")
    
    print(f"Using default: {key} = {default}")
    return default

if __name__ == "__main__":
    print("=== Configuration Test ===")
    
    model_type = get_config('MODEL_TYPE', 'openai')
    print(f"Model Type: {model_type}")
    
    api_key = get_config('OPENAI_API_KEY')
    print(f"API Key loaded: {'Yes' if api_key else 'No'}")
    
    if api_key:
        masked_key = api_key[:10] + "***" + api_key[-4:] if len(api_key) > 14 else "***"
        print(f"Masked key: {masked_key}")