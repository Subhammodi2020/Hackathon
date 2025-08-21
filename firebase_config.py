import firebase_admin
from firebase_admin import credentials, firestore
import os

def initialize_firebase():
    try:
        # Check if Firebase app is already initialized
        if not firebase_admin._apps:
            cred_path = os.path.join(os.path.dirname(__file__), 'firebase_private_key.json')
            if not os.path.exists(cred_path):
                raise FileNotFoundError(f"Firebase credentials file not found at: {cred_path}")
            
            cred = credentials.Certificate(cred_path)
            firebase_admin.initialize_app(cred)
            print("Firebase initialized successfully")
        return firestore.client()
    except Exception as e:
        print(f"Error initializing Firebase: {str(e)}")
        raise

# Initialize Firestore client
try:
    db = initialize_firebase()
    print("Firestore client initialized")
except Exception as e:
    print(f"Failed to initialize Firestore: {str(e)}")
    db = None