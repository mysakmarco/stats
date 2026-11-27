// Web Push для скрінера без сторонніх бібліотек (лише node:crypto).
// - VAPID (RFC 8292): підпис ES256 тим самим ключем P-256, публічна частина якого вшита в сторінку.
// - Шифрування повідомлення aes128gcm (RFC 8291 / RFC 8188).
// - Підписки пристроїв лежать у публічному репозиторії зашифрованими (ECDH на той самий публічний ключ +
//   HKDF-SHA256 + AES-256-GCM), тому прочитати їх і слати сповіщення може лише бот із секретом.
import crypto from "node:crypto";

const b64u = b => Buffer.from(b).toString("base64url");
const unb64u = s => Buffer.from(String(s), "base64url");
const SUB_INFO = Buffer.from("screener-sub-v1");

// ключ: base64url 32 байти приватного скаляра P-256 (секрет SCREENER_PUSH_KEY)
export function loadKey(d64) {
  const d = unb64u((d64 || "").trim());
  if (d.length !== 32) throw new Error("SCREENER_PUSH_KEY має бути 32 байти в base64url");
  const ecdh = crypto.createECDH("prime256v1");
  ecdh.setPrivateKey(d);
  const pub = ecdh.getPublicKey(); // 65 байт, 04||x||y
  const priv = crypto.createPrivateKey({
    key: {kty: "EC", crv: "P-256", d: b64u(d), x: b64u(pub.subarray(1, 33)), y: b64u(pub.subarray(33))},
    format: "jwk",
  });
  return {d, pub, pub64: b64u(pub), priv, ecdh};
}

// розшифрувати запис підписки з subs.json: {e, iv, ct} (зашифровано на сторінці WebCrypto)
export function openSub(key, rec) {
  const shared = key.ecdh.computeSecret(unb64u(rec.e));
  const k = Buffer.from(crypto.hkdfSync("sha256", shared, Buffer.alloc(0), SUB_INFO, 32));
  const ct = unb64u(rec.ct), iv = unb64u(rec.iv);
  const dc = crypto.createDecipheriv("aes-256-gcm", k, iv);
  dc.setAuthTag(ct.subarray(ct.length - 16));
  const pt = Buffer.concat([dc.update(ct.subarray(0, ct.length - 16)), dc.final()]);
  return JSON.parse(pt.toString("utf8"));
}

// та сама процедура, що й на сторінці (для тестів і на випадок, якщо знадобиться з Node)
export function sealSub(pub, obj) {
  const eph = crypto.createECDH("prime256v1");
  eph.generateKeys();
  const shared = eph.computeSecret(pub);
  const k = Buffer.from(crypto.hkdfSync("sha256", shared, Buffer.alloc(0), SUB_INFO, 32));
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv("aes-256-gcm", k, iv);
  const ct = Buffer.concat([c.update(Buffer.from(JSON.stringify(obj))), c.final(), c.getAuthTag()]);
  return {e: b64u(eph.getPublicKey()), iv: b64u(iv), ct: b64u(ct)};
}

function vapidJWT(key, aud, sub) {
  const head = b64u(JSON.stringify({typ: "JWT", alg: "ES256"}));
  const body = b64u(JSON.stringify({aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub}));
  const sig = crypto.sign("sha256", Buffer.from(head + "." + body), {key: key.priv, dsaEncoding: "ieee-p1363"});
  return head + "." + body + "." + b64u(sig);
}

// RFC 8291: шифрування тіла під ключі браузера (p256dh, auth)
export function encrypt(payload, p256dh, auth, _test = {}) {
  const ua = unb64u(p256dh), authB = unb64u(auth);
  if (ua.length !== 65 || authB.length !== 16) throw new Error("bad subscription keys");
  const as = crypto.createECDH("prime256v1");
  if (_test.asPriv) as.setPrivateKey(unb64u(_test.asPriv)); else as.generateKeys();
  const asPub = as.getPublicKey();
  const shared = as.computeSecret(ua);
  const hmac = (k, d) => crypto.createHmac("sha256", k).update(d).digest();
  const prkKey = hmac(authB, shared);
  const ikm = hmac(prkKey, Buffer.concat([Buffer.from("WebPush: info\0"), ua, asPub, Buffer.from([1])]));
  const salt = _test.salt ? unb64u(_test.salt) : crypto.randomBytes(16);
  const prk = hmac(salt, ikm);
  const cek = hmac(prk, Buffer.from("Content-Encoding: aes128gcm\0\x01", "binary")).subarray(0, 16);
  const nonce = hmac(prk, Buffer.from("Content-Encoding: nonce\0\x01", "binary")).subarray(0, 12);
  const c = crypto.createCipheriv("aes-128-gcm", cek, nonce);
  const data = Buffer.concat([Buffer.from(payload, "utf8"), Buffer.from([2])]);
  const ct = Buffer.concat([c.update(data), c.final(), c.getAuthTag()]);
  const rs = Buffer.alloc(4);
  rs.writeUInt32BE(4096);
  return Buffer.concat([salt, rs, Buffer.from([65]), asPub, ct]);
}

// надіслати одне сповіщення; повертає HTTP-статус (201 = прийнято, 404/410 = підписка мертва)
export async function send(key, sub, payload, {ttl = 3 * 3600, urgency = "high", topic, fetchFn = fetch, contact} = {}) {
  const url = new URL(sub.endpoint);
  const jwt = vapidJWT(key, url.origin, contact || "https://github.com/mysakmarco/stats");
  const body = encrypt(typeof payload === "string" ? payload : JSON.stringify(payload), sub.keys.p256dh, sub.keys.auth);
  const headers = {
    "Content-Type": "application/octet-stream",
    "Content-Encoding": "aes128gcm",
    TTL: String(ttl),
    Urgency: urgency,
    Authorization: `vapid t=${jwt}, k=${key.pub64}`,
  };
  if (topic) headers.Topic = String(topic).replace(/[^A-Za-z0-9_-]/g, "").slice(0, 32);
  const ctl = new AbortController(), tm = setTimeout(() => ctl.abort(), 15000);
  try {
    const r = await fetchFn(sub.endpoint, {method: "POST", headers, body, signal: ctl.signal});
    return r.status;
  } finally { clearTimeout(tm); }
}
