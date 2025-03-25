#!/bin/bash
set -e

echo "Deploying HTML files.."
chown -R nginx:nginx /usr/share/nginx/html
chmod -R 755 /usr/share/nginx/html

echo "Restarting Nginx..."
systemctl restart nginx

echo "Deployment completed!"
