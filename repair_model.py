from sentence_transformers import SentenceTransformer
import os

# Use the exact folder you already created
model_path = os.path.abspath("./local_model")

print("Repairing model structure...")
# This downloads the missing config.json and 1_Pooling files
model = SentenceTransformer('all-MiniLM-L6-v2')
model.save(model_path)

print(f"✅ Repair Complete! Files saved to: {model_path}")