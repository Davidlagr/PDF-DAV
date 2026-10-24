from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID
from cryptography import x509
from cryptography.hazmat.primitives import hashes
import datetime

# ==========================================
# 1. CONFIGURACIÓN DE TU FIRMA
# ==========================================
# Esta es la contraseña EXACTA que te pedirá la aplicación Web
CONTRASENA_APP = b"1808DV" 
NOMBRE_ARCHIVO = "mi_firma_limpia.p12"

print("1. Generando clave privada segura (RSA 2048)...")
private_key = rsa.generate_private_key(
    public_exponent=65537,
    key_size=2048,
)

print("2. Configurando detalles legales del certificado...")
# Estos datos quedarán incrustados permanentemente en tu firma PDF
subject = issuer = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, u"David Moisés Lagos Riaño"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, u"Abogado Especialista"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, u"CO"),
])

print("3. Firmando certificado (Validez: 5 años)...")
cert = x509.CertificateBuilder().subject_name(
    subject
).issuer_name(
    issuer
).public_key(
    private_key.public_key()
).serial_number(
    x509.random_serial_number()
).not_valid_before(
    datetime.datetime.utcnow()
).not_valid_after(
    datetime.datetime.utcnow() + datetime.timedelta(days=1825)
).sign(private_key, hashes.SHA256())

print("4. Empaquetando y cifrando el archivo .p12...")
p12_bytes = serialization.serialize_key_and_certificates(
    b"Firma Digital David", # Nombre interno (no es la clave)
    private_key,
    cert,
    None,
    serialization.BestAvailableEncryption(CONTRASENA_APP)
)

print("5. Guardando archivo en el disco...")
with open(NOMBRE_ARCHIVO, "wb") as f:
    f.write(p12_bytes)

print("\n" + "="*50)
print(f"✅ ¡ÉXITO! Se ha creado el archivo: {NOMBRE_ARCHIVO}")
print(f"🔑 La contraseña que debes usar en la web es: 1808DV")
print("="*50 + "\n")
