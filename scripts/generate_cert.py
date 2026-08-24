# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
#
# ADVERTENCIA LEGAL (MÉXICO Y GLOBAL):
# Este código fuente y su arquitectura son propiedad intelectual exclusiva de
# JAVIER ILLAN GONZALEZ. Queda estrictamente prohibida su reproducción,
# distribución, modificación, ingeniería inversa, copia o uso comercial sin la
# autorización expresa y por escrito del autor. Obra protegida conforme a la
# Ley Federal del Derecho de Autor y tratados internacionales aplicables.
# ============================================================================
# -*- coding: utf-8 -*-
import os, datetime
import ipaddress
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CERTS_DIR = os.path.join(BASE_DIR, 'certs')
os.makedirs(CERTS_DIR, exist_ok=True)

cert_path = os.path.join(CERTS_DIR, 'cert.pem')
key_path = os.path.join(CERTS_DIR, 'key.pem')

def ensure_ssl_certs():
    if os.path.exists(cert_path) and os.path.exists(key_path):
        return cert_path, key_path

    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    with open(key_path, 'wb') as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))

    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, 'MX'),
        x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, 'Jalisco'),
        x509.NameAttribute(NameOID.LOCALITY_NAME, 'Guadalajara'),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, 'Orange Crew / O-Leasing'),
        x509.NameAttribute(NameOID.COMMON_NAME, 'localhost'),
    ])

    san = x509.SubjectAlternativeName([
        x509.DNSName('localhost'),
        x509.DNSName('127.0.0.1'),
        x509.IPAddress(ipaddress.IPv4Address('127.0.0.1')),
    ])

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.now(datetime.timezone.utc))
        .not_valid_after(datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=3650))
        .add_extension(san, critical=False)
        .sign(key, hashes.SHA256())
    )

    with open(cert_path, 'wb') as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

    return cert_path, key_path

if __name__ == '__main__':
    cp, kp = ensure_ssl_certs()
    print(f'SSL Certs verified at:\n- {cp}\n- {kp}')
