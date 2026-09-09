FROM airbyte/python-connector-base:3.0.0

WORKDIR /airbyte/integration_code

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py ./
COPY destination_dust ./destination_dust

ENV AIRBYTE_ENTRYPOINT="python /airbyte/integration_code/main.py"
ENTRYPOINT ["python", "/airbyte/integration_code/main.py"]

LABEL io.airbyte.name=airbyte/destination-dust
LABEL io.airbyte.version=0.1.0
