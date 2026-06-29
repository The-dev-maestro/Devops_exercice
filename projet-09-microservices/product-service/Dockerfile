FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm ci --only=production
EXPOSE 3000
USER node
CMD ["node", "server.js"]