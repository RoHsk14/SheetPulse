# Dockerfile for ws-sheets-bot (Render deployment)
# Base image with Node.js and Debian package manager
FROM node:20-slim

# Install Google Chrome (stable) and required utilities
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        ca-certificates wget gnupg && \
    wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | apt-key add - && \
    echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && \
    apt-get install -y --no-install-recommends google-chrome-stable && \
    apt-get purge -y --auto-remove wget gnupg && \
    rm -rf /var/lib/apt/lists/*

# Set Puppeteer to use the Chrome binary we installed, skip its own download
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/google-chrome-stable
ENV PUPPETEER_SKIP_DOWNLOAD=1
ENV NODE_ENV=production

# Application work directory
WORKDIR /app

# Copy package files and install dependencies
COPY package*.json ./
RUN npm ci

# Copy the rest of the source code
COPY . .

# Expose the default port (3000) – can be overridden via $PORT
EXPOSE 3000

# Start the bot
CMD ["node", "server.js"]
