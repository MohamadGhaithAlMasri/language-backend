const crypto = require('crypto');

// SHA-256 to create a secure 32-byte key from our string
const SECRET_SEED = 'MOI_Language_Exam_Secure_Key_26!';
const KEY = crypto.createHash('sha256').update(SECRET_SEED).digest();

function encrypt(plainText) {
  try {
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv('aes-256-cbc', KEY, iv);
    let encrypted = cipher.update(plainText, 'utf8', 'base64');
    encrypted += cipher.final('base64');
    
    // Combine IV and Ciphertext in format: "iv:ciphertext"
    const combined = iv.toString('base64') + ':' + encrypted;
    // Base64 encode the combined result
    return Buffer.from(combined, 'utf8').toString('base64');
  } catch (error) {
    console.error('Encryption error:', error);
    return null;
  }
}

function decrypt(encryptedBase64) {
  try {
    const combined = Buffer.from(encryptedBase64, 'base64').toString('utf8');
    const parts = combined.split(':');
    if (parts.length !== 2) return null;
    
    const iv = Buffer.from(parts[0], 'base64');
    const ciphertextBase64 = parts[1];
    
    const decipher = crypto.createDecipheriv('aes-256-cbc', KEY, iv);
    let decrypted = decipher.update(ciphertextBase64, 'base64', 'utf8');
    decrypted += decipher.final('utf8');
    return decrypted;
  } catch (error) {
    console.error('Decryption error:', error);
    return null;
  }
}

module.exports = {
  encrypt,
  decrypt
};
