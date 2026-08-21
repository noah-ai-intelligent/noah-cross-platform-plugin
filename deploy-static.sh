#!/bin/bash
set -e

echo "Building the static site..."
npm install
npm run build

echo "Deploying to /var/www/html/noah-office..."
cp -r dist/* /var/www/html/noah-office/
cp manifest.xml /var/www/html/noah-office/

echo "Deployment complete! Served at https://noah-office.enpointe.io"
