FROM python:3.9.1-slim-buster 

# turns off buffer for easier container logging
ENV PYTHONUNBUFFERED 1


WORKDIR /app

COPY Excel_crud.py /app/Excel_crud.py 

# copying dependencies
COPY requirements.txt .

# Excel file on which operation is performed
COPY Employee_details.xlsx .

# installing dependencies
RUN pip install -r requirements.txt

# run this command after installing dependencies
CMD [ "python", "./Excel_crud.py" ]