#!/bin/bash
set -e

echo "Building React app for Google Apps Script..."
npm run build:google

echo "Patching index.html to avoid type='module' which Apps Script ignores..."
if [ "$(uname)" = "Darwin" ]; then
  sed -i '' 's/type="module" crossorigin/type="text\/javascript"/g' dist/src/taskpane/index.html
  sed -i '' '/office\.js/d' dist/src/taskpane/index.html
else
  sed -i 's/type="module" crossorigin/type="text\/javascript"/g' dist/src/taskpane/index.html
  sed -i '/office\.js/d' dist/src/taskpane/index.html
fi

echo "Copying build artifacts..."
cp dist/src/taskpane/index.html google/index.html

echo "Pushing code to Apps Script..."
cd google
npx clasp push --force

echo "Deploying new version..."
npx clasp deploy --description "Auto-deployed via deploy.sh"

echo "Deployment complete!"
