# Static hosting for the Office add-in: manifest.xml + the built taskpane/
# login/commands bundles, served over HTTPS behind the existing ingress
# (Office requires TLS for every hosted page, not just the dev server).
# Same build→nginx pattern as noah-frontend-v2/Dockerfile.

# ---- build stage ----
FROM node:20-alpine AS build
WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci

COPY . .
RUN npm run build

# ---- serve stage ----
FROM nginx:1.27-alpine AS serve
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /app/dist /usr/share/nginx/html
# manifest.xml is fetched directly (Integrated Apps / AppSource point at its
# raw URL), so it ships alongside the built bundle, not through Vite.
COPY manifest.xml /usr/share/nginx/html/manifest.xml
EXPOSE 3000
