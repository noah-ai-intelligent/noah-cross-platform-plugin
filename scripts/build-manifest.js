import fs from 'fs';
import path from 'path';

// Parse simple .env files without dependencies
function parseEnvFile(filePath) {
  const env = {};
  if (fs.existsSync(filePath)) {
    const contents = fs.readFileSync(filePath, 'utf-8');
    for (const line of contents.split('\n')) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) continue;
      const [key, ...values] = trimmed.split('=');
      if (key) env[key.trim()] = values.join('=').trim();
    }
  }
  return env;
}

const envName = process.argv[2] || 'development';
const envPath = path.resolve(process.cwd(), `.env.${envName}`);
console.log(`Building manifest for ${envName} using ${envPath}`);

const envVars = { ...parseEnvFile(envPath), ...process.env };

const templatePath = path.resolve(process.cwd(), 'manifest.template.xml');
const outputPath = path.resolve(process.cwd(), 'manifest.xml');

if (!fs.existsSync(templatePath)) {
  console.error(`Error: ${templatePath} not found.`);
  process.exit(1);
}

let manifest = fs.readFileSync(templatePath, 'utf-8');

// Replace placeholders
const noahOfficeUrl = envVars.VITE_NOAH_OFFICE_URL || 'https://localhost:5175';
const noahApiUrl = envVars.VITE_NOAH_API_URL || 'https://noah.enpointe.io/api/v1';

// In manifest, NOAH_API_URL should actually point to the base domain
const noahApiBase = new URL(noahApiUrl).origin;

manifest = manifest.replace(/\{\{NOAH_OFFICE_URL\}\}/g, noahOfficeUrl);
manifest = manifest.replace(/\{\{NOAH_API_URL\}\}/g, noahApiBase);

fs.writeFileSync(outputPath, manifest);
console.log(`Successfully generated manifest.xml with NOAH_OFFICE_URL=${noahOfficeUrl} and NOAH_API_URL=${noahApiBase}`);
