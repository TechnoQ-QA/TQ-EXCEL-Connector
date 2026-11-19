"""TQ Excel Connector - RSA, certificates and licences, with the standard library only.

- make_certificate(): a new key and a self-signed certificate for https://localhost, made on each
  PC at install time, so no PC shares a private key with another. Written as PEM for the server.
- sign() / verify(): RSA PKCS#1 v1.5 with SHA-256, used for licence files. Techno Q keeps the
  private key; the connector holds only the public key (licence_key.py).

Plain Python integers are enough for this: Python's pow() does the modular arithmetic.
"""

import base64
import datetime
import hashlib
import secrets

SMALL_PRIMES = [p for p in range(3, 2000, 2) if all(p % d for d in range(3, int(p ** 0.5) + 1, 2))]
SHA256_PREFIX = bytes.fromhex("3031300d060960864801650304020105000420")


# --------------------------------------------------------------------- RSA --

def _probable_prime(n, rounds=40):
    if n < 2:
        return False
    for p in SMALL_PRIMES:
        if n % p == 0:
            return n == p
    d, r = n - 1, 0
    while d % 2 == 0:
        d //= 2
        r += 1
    for _ in range(rounds):
        x = pow(secrets.randbelow(n - 3) + 2, d, n)
        if x in (1, n - 1):
            continue
        for _ in range(r - 1):
            x = pow(x, 2, n)
            if x == n - 1:
                break
        else:
            return False
    return True


def _prime(bits, e):
    while True:
        candidate = secrets.randbits(bits) | (3 << (bits - 2)) | 1   # top two bits set: n has the full length
        if candidate % e != 1 and _probable_prime(candidate):
            return candidate


def generate_key(bits=2048, e=65537):
    """A new RSA key as a dict of integers: n, e, d, p, q, dp, dq, qinv."""
    while True:
        p, q = _prime(bits // 2, e), _prime(bits // 2, e)
        if p == q:
            continue
        n = p * q
        if n.bit_length() != bits:
            continue
        phi = (p - 1) * (q - 1)
        d = pow(e, -1, phi)
        if p < q:
            p, q = q, p
        return {"n": n, "e": e, "d": d, "p": p, "q": q, "dp": d % (p - 1), "dq": d % (q - 1), "qinv": pow(q, -1, p)}


def _size(n):
    return (n.bit_length() + 7) // 8


def _encoded(message, size):
    """EMSA-PKCS1-v1_5 for SHA-256."""
    digest = SHA256_PREFIX + hashlib.sha256(message).digest()
    if size < len(digest) + 11:
        raise ValueError("RSA key too short")
    return b"\x00\x01" + b"\xff" * (size - len(digest) - 3) + b"\x00" + digest


def sign(key, message):
    size = _size(key["n"])
    m = int.from_bytes(_encoded(message, size), "big")
    s1, s2 = pow(m, key["dp"], key["p"]), pow(m, key["dq"], key["q"])   # Chinese remainder: four times faster
    h = (key["qinv"] * (s1 - s2)) % key["p"]
    s = s2 + h * key["q"]
    if pow(s, key["e"], key["n"]) != m:
        raise ValueError("RSA signature check failed")
    return s.to_bytes(size, "big")


def verify(public, message, signature):
    n, e = public["n"], public["e"]
    size = _size(n)
    if len(signature) != size:
        return False
    s = int.from_bytes(signature, "big")
    if s >= n:
        return False
    return secrets.compare_digest(pow(s, e, n).to_bytes(size, "big"), _encoded(message, size))


# --------------------------------------------------------------------- DER --

def _der(tag, content):
    length = len(content)
    if length < 0x80:
        head = bytes([length])
    else:
        raw = length.to_bytes((length.bit_length() + 7) // 8, "big")
        head = bytes([0x80 | len(raw)]) + raw
    return bytes([tag]) + head + content


def _int(value):
    raw = value.to_bytes(max(1, (value.bit_length() + 8) // 8), "big")   # room for a leading 0 when the high bit is set
    return _der(0x02, raw)


def _seq(*parts):
    return _der(0x30, b"".join(parts))


def _oid(dotted):
    numbers = [int(x) for x in dotted.split(".")]
    out = bytes([40 * numbers[0] + numbers[1]])
    for number in numbers[2:]:
        chunk = [number & 0x7F]
        number >>= 7
        while number:
            chunk.insert(0, 0x80 | (number & 0x7F))
            number >>= 7
        out += bytes(chunk)
    return _der(0x06, out)


def _bits(content):
    return _der(0x03, b"\x00" + content)


def _octets(content):
    return _der(0x04, content)


def _time(moment):
    if moment.year < 2050:
        return _der(0x17, moment.strftime("%y%m%d%H%M%SZ").encode())
    return _der(0x18, moment.strftime("%Y%m%d%H%M%SZ").encode())


def _pem(label, der):
    body = base64.b64encode(der).decode()
    lines = [body[i:i + 64] for i in range(0, len(body), 64)]
    return "-----BEGIN %s-----\n%s\n-----END %s-----\n" % (label, "\n".join(lines), label)


def private_key_pem(key):
    """PKCS#1 "RSA PRIVATE KEY", which Python's ssl module loads."""
    return _pem("RSA PRIVATE KEY", _seq(*[_int(x) for x in (0, key["n"], key["e"], key["d"], key["p"], key["q"],
                                                             key["dp"], key["dq"], key["qinv"])]))


def make_certificate(host="localhost", days=825, key=None):
    """Returns (certificate PEM, private key PEM, SHA-1 thumbprint) for a server certificate for host.

    Self-signed, not a certificate authority (it cannot be used to vouch for any other site), valid
    for host and 127.0.0.1, for server authentication only.
    """
    key = key or generate_key()
    sha256_rsa = _seq(_oid("1.2.840.113549.1.1.11"), _der(0x05, b""))
    name = _seq(_der(0x31, _seq(_oid("2.5.4.3"), _der(0x0C, host.encode()))),
                _der(0x31, _seq(_oid("2.5.4.10"), _der(0x0C, b"TQ Connector (this PC only)"))))
    public = _seq(_int(key["n"]), _int(key["e"]))
    spki = _seq(_seq(_oid("1.2.840.113549.1.1.1"), _der(0x05, b"")), _bits(public))
    now = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)

    def extension(oid, value, critical=False):
        return _seq(_oid(oid), *([_der(0x01, b"\xff")] if critical else []), _octets(value))

    san = _seq(_der(0x82, host.encode()), _der(0x87, bytes([127, 0, 0, 1])))
    extensions = _seq(
        extension("2.5.29.19", _seq(), critical=True),                       # basic constraints: not a CA
        extension("2.5.29.15", _der(0x03, b"\x05\xa0"), critical=True),      # key usage: signature, key encipherment
        extension("2.5.29.37", _seq(_oid("1.3.6.1.5.5.7.3.1"))),            # extended key usage: server auth
        extension("2.5.29.17", san),                                         # subject alternative names
        extension("2.5.29.14", _octets(hashlib.sha1(public).digest())),     # subject key identifier
    )
    tbs = _seq(
        _der(0xA0, _int(2)),                                                  # version 3
        _int(secrets.randbits(120) | (1 << 120)),
        sha256_rsa,
        name,
        _seq(_time(now - datetime.timedelta(days=1)), _time(now + datetime.timedelta(days=days))),
        name,
        spki,
        _der(0xA3, extensions),
    )
    certificate = _seq(tbs, sha256_rsa, _bits(sign(key, tbs)))
    return _pem("CERTIFICATE", certificate), private_key_pem(key), hashlib.sha1(certificate).hexdigest().upper()
