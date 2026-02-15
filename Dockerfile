FROM node:20-slim
WORKDIR /app
COPY package*.json ./
RUN npm ci --only=production
COPY . .
RUN mkdir -p uploads exports
ENV NODE_ENV=production
ENV PORT=8080
EXPOSE 8080
CMD ["node", "backend/server.js"]
