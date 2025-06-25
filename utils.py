# utils.py
import uuid
import hashlib

def obtener_id_maquina():
    id_unico = str(uuid.getnode())
    hash_id = hashlib.sha256(id_unico.encode()).hexdigest()
    return hash_id
