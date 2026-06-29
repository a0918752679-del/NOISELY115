FROM node:20-alpine
WORKDIR /app
ENV NODE_ENV=production
COPY package.json ./
RUN npm install --omit=dev --no-audit --no-fund --registry=https://registry.npmjs.org/
COPY . .
RUN mkdir -p data
EXPOSE 8080
CMD ["npm", "start"]
