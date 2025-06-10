import os
import base64
import io
import json
import logging
import pandas as pd
from flask import Flask, render_template, request, flash, redirect, url_for, Response, session
from flask_session import Session
from sqlalchemy import create_engine, Column, String, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")

# Monkey-patch for Flask-Session in Flask 2.3+
if not hasattr(app, 'session_cookie_name'):
    app.session_cookie_name = app.config.get("SESSION_COOKIE_NAME", "session")

app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_FILE_DIR"] = "./.flask_session/"
app.config["SESSION_PERMANENT"] = False
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
Session(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Configure SQLAlchemy using Render's external database URL.
# It is best to store the connection string in an environment variable.
PG_DATABASE_URL = os.environ.get("PG_DATABASE_URL", 
    "postgresql://render_postgres_db_7cik_user:123456@dpg-d1479obipnbc73c4hts0-a.oregon-postgres.render.com/render_postgres_db_7cik"
)

engine = create_engine(PG_DATABASE_URL)
Base = declarative_base()

class Country(Base):
    __tablename__ = 'countries'
    iso_code = Column(String(3), primary_key=True)
    country_name = Column(Text, nullable=False)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# A helper function that performs a lookup for the country name based on the first three letters of a TADIG code.
def get_country_from_tadig(tadig_code):
    lookup_code = tadig_code[:3].upper()
    db = SessionLocal()
    try:
        country_obj = db.query(Country).filter(Country.iso_code == lookup_code).first()
        if country_obj:
            return country_obj.country_name
        else:
            return "Unknown"
    except Exception as e:
        logger.exception("Error looking up country for %s", lookup_code)
        return "Error"
    finally:
        db.close()

# Example Excel processing route that integrates the lookup:
@app.route("/process_excel", methods=["POST"])
def process_excel():
    file = request.files.get("file")
    if file and file.filename.endswith((".xlsx", ".xls")):
        try:
            file_bytes = file.read()
            df = pd.read_excel(io.BytesIO(file_bytes), header=0)
            # Assuming your Excel file has a column "TADIG PLMN Code"
            # Append a 'Country' column by looking up based on first 3 letters.
            df["Country"] = df["TADIG PLMN Code"].apply(get_country_from_tadig)
            
            # Optional: Convert the dataframe data to HTML for rendering.
            table_html = df.to_html(classes="data-table", index=False)
            return render_template("results.html", table_html=table_html)
        except Exception as e:
            logger.exception("Error processing Excel file")
            flash(f"Error processing Excel file: {e}", "error")
    else:
        flash("Please upload a valid Excel file.", "error")
    return redirect(url_for("index"))

@app.route("/")
def index():
    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.debug("Starting Flask server on port %d", port)
    app.run(debug=True, host="0.0.0.0", port=port)
