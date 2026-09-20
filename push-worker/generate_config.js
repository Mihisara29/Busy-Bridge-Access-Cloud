const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const webpush = require('web-push');

const subjectArg = process.argv[2] || 'mailto:admin@example.com';
const subject = /^mailto:|^https:\/\//i.test(subjectArg)
  ? subjectArg
  : `mailto:${subjectArg}`;

const outputPath = path.resolve(
  __dirname,
  '..',
  'data',
  'push',
  'push_config.json',
);

fs.mkdirSync(path.dirname(outputPath), { recursive: true });

if (fs.existsSync(outputPath)) {
  console.log(`Push config already exists: ${outputPath}`);
  console.log('Existing VAPID keys were preserved.');
  process.exit(0);
}

const vapid = webpush.generateVAPIDKeys();

const config = {
  publicKey: vapid.publicKey,
  privateKey: vapid.privateKey,
  subject,
  workerSecret: crypto.randomBytes(32).toString('hex'),
  createdAt: new Date().toISOString(),
};

fs.writeFileSync(
  outputPath,
  JSON.stringify(config, null, 2),
  'utf8',
);

console.log(`Created BUSY Cloud push config: ${outputPath}`);
console.log(`VAPID subject: ${subject}`);
console.log('Back up this file. If VAPID keys are changed, browsers must subscribe again.');
