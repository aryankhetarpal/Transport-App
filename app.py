from flask import Flask, render_template, request
import pandas as pd
import re

app = Flask(__name__)

# ---------- Load Excel ----------
excel_file = r"C:\Users\Aryan\Desktop\Transport App\TRANSPORT CHARGES._.xlsx"
sheet = pd.read_excel(excel_file, sheet_name="JAIDURGA LOGISTIC", header=None)

records = []
current_transporter = None
headers = None
previous_weight_row = None

# ---------- Helper: Check if a row is a valid transporter ----------
def is_valid_transporter(name):
    name_upper = name.upper()
    # Must have at least 2 words
    if len(name.split()) < 2:
        return False
    # Must only contain letters, spaces, &, -, /
    if not all(c.isalpha() or c.isspace() or c in "&-/" for c in name):
        return False
    # Must not contain digits
    if any(char.isdigit() for char in name):
        return False
    # Must not contain common non-transporter keywords
    blacklist = ["LOCATION", "TON", "CHARGES", "EXTRA", "KG", "KM", "AMOUNT"]
    if any(keyword in name_upper for keyword in blacklist):
        return False
    return True

# ---------- Parse Excel ----------
for i, row in sheet.iterrows():
    first_cell = str(row[1] if pd.notna(row[1]) else row[0]).strip()

    # ---------- Detect Transporter ----------
    if first_cell and "WEIGHT" not in first_cell.upper() and "CAPACITY" not in first_cell.upper():
        if is_valid_transporter(first_cell):
            current_transporter = first_cell
            headers = None
            previous_weight_row = None
            print(f"\nDetected transporter: {current_transporter}")
            continue

    # ---------- Detect Split Header (Weight + Capacity) ----------
    if re.search(r"^WEIGHT$", first_cell.upper()):
        previous_weight_row = i
        continue

    if "CAPACITY" in first_cell.upper() and previous_weight_row is not None:
        prev_row = sheet.iloc[previous_weight_row]
        combined = [
            (str(a).strip() if pd.notna(a) else "") + " " + (str(b).strip() if pd.notna(b) else "")
            for a, b in zip(prev_row[1:], row[1:])
        ]
        headers = [h.strip() if h.strip() else f"Loc_{j}" for j, h in enumerate(combined, start=2)]
        previous_weight_row = None
        continue

    # ---------- Normal Header Detection ----------
    if re.search(r"WEIGHT.*CAPACITY|CAPACITY.*WEIGHT|WEIGHT\s*\(.*\)", first_cell.upper()):
        headers = []
        for j, c in enumerate(row[1:], start=2):
            if pd.notna(c):
                headers.append(str(c).strip())
            else:
                headers.append(f"Loc_{j}")
        continue

    # ---------- Parse Rate Rows ----------
    if current_transporter and headers and pd.notna(row[1]) and "WEIGHT" not in first_cell.upper():
        weight_label = str(row[1]).strip()
        for j, loc_header in enumerate(headers[1:], start=2):
            if j < len(row):
                val = row[j]
                val_str = str(val)
                val_clean = re.sub(r"[^\d.]", "", val_str)
                try:
                    rate = float(val_clean) if val_clean else None
                except ValueError:
                    rate = None
                if rate is not None:
                    # Split combined locations and create a record for each
                    split_locations = [l.strip().title() for l in loc_header.split('/') if l.strip()]
                    for loc in split_locations:
                        records.append({
                            "Transporter": current_transporter,
                            "WeightLabel": weight_label,
                            "Location": loc,
                            "Rate": rate
                        })

# ---------- Create DataFrame ----------
df = pd.DataFrame(records)

# ---------- Flask Routes ----------
@app.route("/")
def index():
    weights = sorted(df["WeightLabel"].unique())
    locations = sorted(df["Location"].unique())
    return render_template("index.html", weights=weights, locations=locations)

@app.route("/get_recommendations", methods=["POST"])
def get_recommendations():
    weight = request.form.get("weight")
    location = request.form.get("location")

    filtered = df[
        (df["WeightLabel"].str.lower() == weight.lower()) &
        (df["Location"].str.lower() == location.lower()) &
        (df["Rate"].notna())
    ]

    if not filtered.empty:
        results = filtered.sort_values("Rate").head(3)
        table_html = """
        <h3>Top 3 Cheapest Options</h3>
        <table border="1" cellpadding="5">
            <tr><th>Transporter</th><th>Weight</th><th>Rate</th></tr>
        """
        for _, row in results.iterrows():
            table_html += f"<tr><td>{row['Transporter']}</td><td>{row['WeightLabel']}</td><td>{row['Rate']}</td></tr>"
        table_html += "</table>"
    else:
        table_html = "<p>No transporters found for that weight and location.</p>"

    return table_html

# ---------- Run Flask ----------
if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
