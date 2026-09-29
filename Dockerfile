FROM python:3.12-slim

WORKDIR /app

# 파이썬 로그가 바로 보이도록
ENV PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=UTF-8

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY erp_ui.py .
COPY .streamlit/config.toml .streamlit/config.toml

# DB는 코드가 아니라 영구 볼륨(/data)에 저장한다. 컨테이너를 새로 배포해도 데이터가 남는다.
ENV ERP_DB_DIR=/data

EXPOSE 8501

CMD ["streamlit", "run", "erp_ui.py", \
     "--server.port=8501", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--browser.gatherUsageStats=false"]
