from pydantic import BaseModel
import openai
from typing import Optional
import base64
import sys
import json

# Configurar la API Key
client = openai.Client(api_key="")

def get_arguments():
    # Comprobar si hay suficientes argumentos
    if len(sys.argv) < 3:
        print("Uso: python example.py <HOJA_EXCEL> <ATRIBUTO1> <ATRIBUTO2> ...")
        sys.exit(1)

    # El primer argumento es la hoja de Excel
    excel_path = sys.argv[1]

    # Los siguientes argumentos son los atributos deseados
    attributes = sys.argv[2:]

    return excel_path, attributes

# Función para crear una clase dinámica de Pydantic
def createDynamicClass(class_name: str, attributes: list):
    # Usamos `__annotations__` para definir los tipos correctamente
    fields = {attr: Optional[str] for attr in attributes}  # Solo anotación de tipo
    defaults = {attr: None for attr in attributes}  # Valores predeterminados
    
    # Crear la clase con `type`
    dynamic_class = type(class_name, (BaseModel,), {"__annotations__": fields, **defaults})
    return dynamic_class

# Function to encode the image
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

# Convierte la respuesta en diccionario y la guarda como JSON
def save_dict_to_json(response_model, outputFile="metadata.json"):
    try:
        # Convertir la respuesta del modelo Pydantic a un diccionario
        response_dict = response_model.model_dump()

        # Guardar el diccionario en un archivo JSON
        with open(outputFile, "w", encoding="utf-8") as json_file:
            json.dump(response_dict, json_file, indent=4, ensure_ascii=False)

        print(f"Datos guardados en {outputFile}")
    except Exception as e:
        print(f"Error al guardar los datos en JSON: {e}")

def programa(attributes):
    # 2. Crear la clase dinámica
    DynamicModel = createDynamicClass("DynamicModel", attributes)

    # Path to your image
    image_path = "images/sh2_cas.png"
    # Getting the base64 string
    base64_image = encode_image(image_path)

    completion = client.beta.chat.completions.parse(
        model="gpt-4o-2024-08-06",
        messages=[
            {"role": "system", "content": "You are an expert at structured data extraction. You will be given an image from an Excel sheet and should convert it into the given structure."},
            {"role": "user", "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{base64_image}"},
                    },
                ],
            }
        ],
        response_format=DynamicModel,
    )

    research_paper = completion.choices[0].message.parsed
    print("Respuesta completa de la API:")
    print(research_paper)

    save_dict_to_json(research_paper)
    

# Ejemplo de uso
if __name__ == "__main__":
    # Obtener los argumentos desde la línea de comandos
    excel_path, attributes = get_arguments()

    print("Atributos solicitados:", attributes)

    programa(attributes)
    
