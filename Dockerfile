
FROM node:20-slim AS release

# Set the working directory
WORKDIR /app

RUN npm install -g excel-mcp-server-pcvelz@0.15.4

# Command to run the application
ENTRYPOINT ["excel-mcp-server"]
