# Asset Return Form - Web Application

A Flask-based web app to fill out the Asset Return Form template.

## Setup

1. Install Python 3.8+

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the app:
```bash
python app.py
```

4. Open your browser and go to:
```
http://localhost:5000
```

## Features

- **All fields optional** — leave blank to skip
- **Clean web UI** — fill forms in your browser
- **Preview** — view the filled Word document
- **Save** — downloads as `Asset_Return_<EmployeeName>_<timestamp>.docx`
- **Print** — sends directly to your default/network printer

## Output

All generated forms are saved in the `Asset_Return_Forms/` folder.

## Changing the Port

Edit `app.py` and change this line:
```python
app.run(host="0.0.0.0", port=5000, debug=True)
```
Replace `5000` with any port you want.
