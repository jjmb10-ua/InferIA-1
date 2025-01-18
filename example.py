import openai
import base64
import ast
import json

# Configurar la API Key
client = openai.Client(api_key="")


# Función para limpiar las etiquetas de formato
def clean_response(response):
    # Eliminar las etiquetas de código Markdown
    if response.startswith("```python"):
        response = response[len("```python"):].strip()
    if response.endswith("```"):
        response = response[:-len("```")].strip()
    
    # Convertir a un diccionario de Python
    try:
        dict = ast.literal_eval(response)  # Alternativa segura a eval()
        print("Diccionario de Python:", dict)
    except Exception as e:
        print(f"Error al procesar la respuesta como diccionario: {e}")
    
    return dict

# Guardar el diccionario como JSON
def save_dict_to_json(data_dict, file_name="metadata.json"):
    with open(file_name, "w", encoding="utf-8") as json_file:
        json.dump(data_dict, json_file, ensure_ascii=False, indent=4)
    print(f"El archivo JSON se ha guardado como: {file_name}")

# Function to encode the image
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

# Path to your image
image_path = "images/sh2_cas.png"
model_gpt = "gpt-4o"

# Getting the base64 string
base64_image = encode_image(image_path)

prompt = """Analyze the content of the following image of Excel sheet. Identify any metadata present in the sheet, such as titles, subtitles, dates, notes, and other relevant details. Return the metadata as a Python dictionary, with keys for each metadata field and their corresponding values.

If no metadata is found, respond with an empty dictionary: {{}}.

Respond strictly with the Python dictionary and no additional text or explanations."""

# Realizar la solicitud
response = client.chat.completions.create(
    model=model_gpt,
    messages=[
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": prompt,
                },
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:image/png;base64,{base64_image}"},
                },
            ],
        }
    ],
)
response_1_content = response.choices[0].message.content
# Imprimir el contenido del mensaje generado
print(response_1_content)

# Limpiar la respuesta
cleaned_response = clean_response(response_1_content)

save_dict_to_json(cleaned_response, "metadata.json")

prompt = """Analyze the content of the following image of Excel sheet. Determine if the sheet contains tabular data. If no tabular data is found, respond with `"No data detected."`.

If tabular data is present, identify the line number where the data starts and extract the column headers. Return the result as a Python dictionary with the following structure:

```python
{
    "Line": <line_number>,  # The line number where the data starts
    "Headers": ["Header1", "Header2", ...]  # List of detected headers
}
"""

# Solicitar la línea de inicio y las cabeceras de los datos
response_3 = client.chat.completions.create(
model=model_gpt,
messages=[
    {"role": "system", "content": "You are a helpful assistant."},
    {
    "role": "user",
    "content": [
        {
            "type": "text",
            "text": prompt,
        },
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{base64_image}"},
            },
    ],
}
]
)
response_3_content = response_3.choices[0].message.content
print("Data details:", response_3_content)
cleaned_response = clean_response(response_3_content)