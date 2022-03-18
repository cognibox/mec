# pull official base image
FROM python:3.9.7-slim

# set work directory
WORKDIR /usr/src/

# set environment variables
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

RUN apt --yes update && apt --yes upgrade && apt-get --yes install netcat

# install dependencies
RUN pip install --no-cache-dir --upgrade pip
COPY ./requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# create directory for the script
RUN mkdir -p /home/script

# create the app user
RUN addgroup --system script && adduser --system --ingroup script script

# create the appropriate directories
ENV HOME=/home/script
WORKDIR $HOME

# copy project
COPY . $HOME
RUN mkdir $HOME/data
RUN chown -R script:script $HOME
VOLUME $HOME/data
# chown all the files to the app user

# change to the app user
USER script
ENTRYPOINT ["python", "main.py"]
