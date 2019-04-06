#!/bin/bash
docker build . -t gcr.io/run-intro/pdf-service
docker run -p 8080:8080 gcr.io/run-intro/pdf-service