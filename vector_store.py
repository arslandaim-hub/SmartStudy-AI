import chromadb
from chromadb.utils import embedding_functions

# Initialize the local database folder
client = chromadb.PersistentClient(path="./study_db")

# Use a lightweight local embedding model
model_name = "all-MiniLM-L6-v2"
embedding_func = embedding_functions.SentenceTransformerEmbeddingFunction(model_name=model_name)

# Create or get a collection for your Zoology/Botany notes
collection = client.get_or_create_collection(
    name="academic_notes", 
    embedding_function=embedding_func
)

def add_notes_to_db(text_chunks, metadata_list, ids):
    """Adds processed text chunks to the vector database."""
    collection.add(
        documents=text_chunks,
        metadatas=metadata_list,
        ids=ids
    )

def query_db(question, n_results=3):
    """Retrieves the most relevant snippets for a question."""
    results = collection.query(
        query_texts=[question],
        n_results=n_results
    )
    return results['documents'][0]