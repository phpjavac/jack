# use docker node 14.10.1
FROM node:14.10.1


# create a directory to run docker
WORKDIR /app

# copy package.json into the new directory
COPY package.json /app

# install the dependencies
RUN npm install

# copy all other files and folder into the app directory
COPY . /app

# open port 5000
EXPOSE 5000

# run the server
CMD node index.js