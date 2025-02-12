from krutrim_cloud import KrutrimCloud
from dotenv import load_dotenv
import os

# Load API key from .env file
load_dotenv()
api_key = os.getenv("KRUTRIM_CLOUD_API_KEY")

# Ensure API key is set
if not api_key:
    raise ValueError("API key not found. Set KRUTRIM_CLOUD_API_KEY in the .env file.")

# Initialize KrutrimCloud client with API key
client = KrutrimCloud(api_key=api_key)

# Define the model and input prompt
model_name = "DeepSeek-R1"
messages = [
    {"role": "user", "content": "How to drain a k8s node?"}
]

try:
    # Make API request
    response = client.chat.completions.create(model=model_name, messages=messages)

    # Access generated output
    txt_output_data = response.choices[0].message.content  # type:ignore
    print(f"Output: \n{txt_output_data}")

    # Optional: Save generated output
    response.save(output_dirpath="./output")

except Exception as exc:
    print(f"Exception: {exc}")