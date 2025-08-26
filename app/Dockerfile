FROM node:23-slim as base
WORKDIR /usr/src/app
COPY package*.json ./

FROM base as dev
RUN npm install -g nodemon
RUN npm install
COPY . .
EXPOSE 3000
CMD ["npm", "run", "start"]

FROM base as production
RUN npm ci --only=production
COPY . .
EXPOSE 3000
CMD ["npm", "run", "start"]
