FROM goreleaser/goreleaser:v2.8.1 AS go-build

# Set the working directory
WORKDIR /app

# Copy files to the working directory
COPY . ./

# Build the project
RUN goreleaser build --snapshot --clean

FROM node:20-slim AS npm-build

# Set the working directory
WORKDIR /app

# Install dependencies and build the project
RUN npm ci && tsc

# Use a smaller Node.js base image for the final stage
FROM node:20-slim AS release

# Set the working directory
WORKDIR /app

# Copy the build output and necessary files from the build stage
COPY --from=go-build  /app/dist /app/dist
COPY --from=npm-build /app/node_modules /app/node_modules
COPY --from=npm-build /app/package.json /app/package.json

# Command to run the application
ENTRYPOINT ["node", "dist/launcher.js"]
