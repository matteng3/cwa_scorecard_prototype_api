
from flask import Flask, request, jsonify, send_file
import pandas as pd

app = Flask(__name__)

# Function to calculate score based on brackets
def calculate_score(value, brackets):
    for bracket in brackets:
        if bracket[0](value):
            return bracket[1]
    return 0

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400

    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    if file:
        xls = pd.ExcelFile(file)
        results = process_metrics(xls)
        results_file = 'metrics_results.csv'
        results.to_csv(results_file, index=False)
        return send_file(results_file, as_attachment=True)

def process_metrics(xls):
    # Load sheets from the Excel file
    channel_mix_df = pd.read_excel(xls, sheet_name='ChannelMix')
    source_traffic_df = pd.read_excel(xls, sheet_name='Source Traffic')
    paid_media_df = pd.read_excel(xls, sheet_name='Paid Media')

    # Calculating the metrics
    results = pd.DataFrame()

    # Metric 1: Channel Mix - Brand.com % of Revenue
    results = results.append(calculate_channel_mix_revenue(channel_mix_df, 'Brand.com', 'Brand.com % of Revenue'), ignore_index=True)
    
    # Metric 2: Channel Mix - OTA % of Revenue
    results = results.append(calculate_channel_mix_revenue(channel_mix_df, 'OTA', 'OTA % of Revenue'), ignore_index=True)

    # Metric 3: Brand.com Conversion % (Bookings / Traffic)
    # Placeholder logic, replace with actual calculation
    results = results.append({'Metric': 'Brand.com Conversion %', 'Result': 0}, ignore_index=True)

    # Metrics 4-10: Placeholder logic, replace with actual calculation
    for i in range(4, 11):
        results = results.append({'Metric': f'Metric {i}', 'Result': 0}, ignore_index=True)

    return results

def calculate_channel_mix_revenue(df, channel_name, metric_name):
    channel_revenue = df[df['s_ChannelName'].str.contains(channel_name, na=False)]['d_Revenue'].sum()
    total_revenue = df['d_Revenue'].sum()
    percentage = (channel_revenue / total_revenue * 100) if total_revenue else 0
    score_brackets = [
        (lambda x: x >= 50, 10),
        (lambda x: 40 <= x < 50, 9),
        (lambda x: 35 <= x < 40, 8),
        (lambda x: 30 <= x < 35, 7),
        (lambda x: x < 30, 6)
    ] if metric_name == 'Brand.com % of Revenue' else [
        (lambda x: x < 5, 10),
        (lambda x: 5 <= x <= 10, 9),
        (lambda x: 11 <= x <= 15, 8),
        (lambda x: 16 <= x <= 20, 7),
        (lambda x: 21 <= x <= 30, 6),
        (lambda x: x > 30, 5)
    ]
    score = calculate_score(percentage, score_brackets)
    return {'Metric': metric_name, 'Result': score}

if __name__ == '__main__':
    app.run(debug=True)
