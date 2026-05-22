/**
 * shared/authMiddleware.js
 *
 * Azure AD tarafından imzalanan Bearer token'ı doğrular ve
 * içindeki kullanıcı kimliğini (oid) çıkarır.
 *
 * Teams SSO token → oid claim = Azure AD Object ID
 *
 * Doğrulama adımları:
 *  1. Token formatı kontrol
 *  2. Azure AD JWKS ile imza doğrulama
 *  3. iss / aud / exp claim doğrulama
 */

const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");

const TENANT_ID  = process.env.TENANT_ID;
const CLIENT_ID  = process.env.CLIENT_ID;

// Azure AD JWKS endpoint'i
const jwks = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`,
  cache: true,
  cacheMaxEntries: 10,
  cacheMaxAge: 60 * 60 * 1000, // 1 saat
});

/** JWKS'ten imzalama anahtarını al */
function getSigningKey(header) {
  return new Promise((resolve, reject) => {
    jwks.getSigningKey(header.kid, (err, key) => {
      if (err) return reject(err);
      resolve(key.getPublicKey());
    });
  });
}

/**
 * HTTP isteğinden Bearer token'ı çıkar ve doğrula.
 * @returns {{ userId: string, email: string, name: string } | null}
 */
async function extractCaller(request) {
  const authHeader =
    request.headers?.get?.("authorization") ||
    request.headers?.authorization ||
    "";

  if (!authHeader.startsWith("Bearer ")) return null;
  const token = authHeader.slice(7);

  try {
    // Header'dan kid al, JWKS'ten public key çek
    const decoded = jwt.decode(token, { complete: true });
    if (!decoded) return null;

    const publicKey = await getSigningKey(decoded.header);

    // İmzayı ve standart claim'leri doğrula
    const payload = jwt.verify(token, publicKey, {
      algorithms: ["RS256"],
      audience: CLIENT_ID,
      issuer: [
        `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
        `https://sts.windows.net/${TENANT_ID}/`,
      ],
    });

    return {
      userId: payload.oid,               // Azure AD Object ID
      email:  payload.upn || payload.preferred_username || payload.email || "",
      name:   payload.name || "",
    };
  } catch (err) {
    // Geliştirme ortamında token doğrulaması atlayılabilir
    if (process.env.SKIP_AUTH_VALIDATION === "true") {
      const raw = jwt.decode(token);
      if (raw?.oid) {
        return {
          userId: raw.oid,
          email:  raw.upn || raw.preferred_username || "",
          name:   raw.name || "",
        };
      }
    }
    console.warn("Token doğrulama başarısız:", err.message);
    return null;
  }
}

/**
 * Görünürlük kurallarını uygular.
 *
 * KURAL:
 *  - beklemede  → sadece talepçi VEYA onaylayıcı görebilir
 *  - onaylandi / reddedildi → talepçi kendi talebini, onaylayıcı hepsini görebilir
 *
 * @param {Array}   requests    - Tüm ham talepler
 * @param {string}  callerId    - İstekte bulunan kullanıcının Azure AD oid'si
 * @param {boolean} isApprover  - Kullanıcı onaylayıcı mı?
 * @returns {Array} Filtrelenmiş talepler
 */
function applyVisibilityRules(requests, callerId, isApprover) {
  return requests.filter((req) => {
    const isOwner = req.requesterId === callerId;

    // Onaylayıcılar her durumda tüm talepleri görür
    if (isApprover) return true;

    // Talep sahibi kendi taleplerinin hepsini görür
    if (isOwner) return true;

    // Diğer kullanıcılar hiçbir talebi göremez
    return false;
  });
}

module.exports = { extractCaller, applyVisibilityRules };
