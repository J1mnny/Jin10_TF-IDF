FROM python:3.6
RUN mkdir /news
WORKDIR /news
RUN pip install --upgrade pip
RUN pip install BeautifulSoup4 openpyxl jieba pandas snownlp xlrd requests schedule

