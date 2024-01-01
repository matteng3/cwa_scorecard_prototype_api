import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
import time






# Load Excel data once and perform date conversion
def load_and_prepare_data(excel_path):
    data = {}
    sheets_to_keys = {
        'channel_mix': 'channel_mix_data',
        'visit_revenue': 'visits_and_rev_data',
        'source_traffic': 'source_traffic_data',
        'paid_media': 'paid_media_data'
    }
    for sheet, key in sheets_to_keys.items():
        df = pd.read_excel(excel_path, sheet_name=sheet)
        df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d')  # Convert date column
        data[key] = df
    return data

excel_path = 'combined_excel_file.xlsx'
data = load_and_prepare_data(excel_path)
print(data.keys())

# Complete lists for Organic and Paid Search Traffic without "(lowercase)"
organic_search_ids = [ '048dba83-e9e3-420b-b2e4-1335ed80524a', '08ecbfb6-4efb-401d-b841-4a18b2524100', '0b83d66d-47cf-497c-99be-ce75fa9c8ac6', '0ee31528-f6f1-4892-85f1-cb8eb68dd916', '18dba0f5-c7b6-4fd8-a184-910a954b8f77', '1bfb776b-3b9c-46f9-b491-bd8afbefd204', '1ce17a21-23ad-44a3-a9cb-d09494ff211d', '1d1135ee-9e00-43bc-8aa9-9dae1d59fe6a', '24dd55e2-beb1-4952-b0c2-b30cf94edd3c', '2692ef59-23ef-4512-b834-3a332d4c0c59', '275e4d28-cb0d-4dbc-9d07-b368a99057ef', '295da3bf-67f4-4a4c-a992-4122d71f27e5', '3394d2c2-9db0-4472-94eb-525c031e5829', '56c416d2-e280-4f0c-875d-7ad43a2cd721', '5b8d6b52-ce50-4168-8d29-51b4edcc4a83', '5cc2bd59-044e-400c-ab86-e41afc0a6b59', '5cc3275c-972f-47bc-9d42-ab04b1378e53', '7cb78729-d8f5-4c5c-9e73-16edf4d51732', '8afd5eca-07e4-4926-b7f8-e0bfba18a03d', 'aaf1577b-fe3b-4b3b-9c54-66934ed9bd96', 'c541e5c7-55c0-4562-95ce-af10bc4bc2d3', 'daf6bd07-a511-4aac-82ad-95f355cc7d61', 'dd905bff-c0c9-4802-a9b6-6eec98ac59ab', 'e2e649be-3eda-4529-8a3d-45de053dc4e7', 'eb4a0eb0-25b4-4256-95c3-c0d6b976ff8f', 'ec30b6b1-9b15-44d4-a42e-fa877fc74de5', 'ed7a896f-d23b-4591-905d-e1ac7115c355', 'ee5db7c5-a6fe-445d-bc09-e923635e6b3b', 'ef221f77-260c-445e-b6dd-8f584861b006', 'ef51d6b0-babb-465b-a996-0c1f8c3b3b0b', 'f831e7c3-9568-409f-b6ec-b099eaa64977']
        # List of UUIDs for paid search sources
paid_search_ids = ['0d1fdf8d-12b9-445e-b1ce-3b6cc56fbbb6', '1ce17a21-23ad-44a3-a9cb-d09494ff211d', '249ccda9-a648-480b-abb0-d583bbf17233', '565c9f42-3e16-4630-a5d3-0b4ec3c77a96', '647f29a9-d427-40fd-ba92-dfd9a8ab6017', 'b7224d5c-bd00-4ea5-925d-54183a0bf20f', 'c84a5595-1fda-468f-8594-a9b0aadef70b', 'd1f5262a-ad63-4edb-b755-629bb4dc587f', 'fd622690-d38d-4355-8ecf-ee912ce35e8f']

# Define the organic and paid metasearch categories based on the image provided.
metasearch_ids = [
    '2692ef59-23ef-4512-b834-3a332d4c0c59',
    '5cc2bd59-044e-400c-ab86-e41afc0a6b59',
    'aaf1577b-fe3b-4b3b-9c54-66934ed9bd96',
    'd1f5262a-ad63-4edb-b755-629bb4dc587f',
    'eb4a0eb0-25b4-4256-95c3-c0d6b976ff8f',
    'fd622690-d38d-4355-8ecf-ee912ce35e8f',
    'b7224d5c-bd00-4ea5-925d-54183a0bf20f',  # Potential non-explicit
    'c84a5595-1fda-468f-8594-a9b0aadef70b'   # Potential non-explicit
]

social_traffic_ids = [
    '0d1fdf8d-12b9-445e-b1ce-3b6cc56fbbb6',  # PAID SOCIAL
    'b7b3d09b-8caa-4904-b188-c5af8c690864',  # ORGANIC SOCIAL
    '7cb78729-d8f5-4c5c-9e73-16edf4d51732',  # NATURAL SOCIAL
    'c541e5c7-55c0-4562-95ce-af10bc4bc2d3',  # SOCIAL MEDIA
]

ota_ids = [
    '3c736439-8c98-4711-af28-edda00114373',
    '46615731-0adc-4bfa-9c86-961939ea1b77',
    '490fea4c-b5ac-49ee-962b-393d60b0aae5',
    '5d72d652-fc44-4cb7-8426-644c6a3988e6',
    '86566284-59e1-4d57-9061-2e4307ece468',
    'c39992b2-d44c-42c3-bc93-ebb723f2a7b3',
    'cdf5f1d4-8bcf-4ed6-8587-54cca950a51b',
]

facebook_id = [
    'f9da8533-b300-4ba2-b9b5-e9c7d2d1f5ab'
]


# Function definitions for calculations


def calculate_brand_com_metrics(data, hotel_id, month_num, year):
     # Extract channel_mix_data DataFrame from the data dictionary
    channel_mix_data = data['channel_mix_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter for the specific month and year
    brand_com_data = channel_mix_data[
        (channel_mix_data['hotel_id'] == hotel_id) & 
        (channel_mix_data['date'].dt.month == month_year.month) & 
        (channel_mix_data['date'].dt.year == month_year.year) & 
        (channel_mix_data['channel_type_id'] == 'a9c5e29d-dbd0-4612-a081-bc0d4afdc88d')
    ]

    # Calculate the total revenue for Brand.com
    brand_com_revenue = brand_com_data['revenue'].sum()

    # Calculate the total revenue for the hotel in the specified month and year
    total_revenue = channel_mix_data[
        (channel_mix_data['hotel_id'] == hotel_id) & 
        (channel_mix_data['date'].dt.month == month_year.month) & 
        (channel_mix_data['date'].dt.year == month_year.year)
    ]['revenue'].sum()

    # Calculate the percentage of revenue from Brand.com
    brand_com_percentage = (brand_com_revenue / total_revenue * 100) if total_revenue else 0
    return brand_com_percentage

def calculate_ota_metrics(data, hotel_id, month_num, year):
     # Extract channel_mix_data DataFrame from the data dictionary
    channel_mix_data = data['channel_mix_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter for OTA revenue for the specific month and year
    ota_data = channel_mix_data[
        (channel_mix_data['hotel_id'] == hotel_id) & 
        (channel_mix_data['date'].dt.month == month_year.month) & 
        (channel_mix_data['date'].dt.year == month_year.year) & 
        (channel_mix_data['channel_type_id'] == '0a062047-6040-420a-86d9-84df582c90ba')
    ]

    # Calculate the total revenue for OTA
    ota_revenue = ota_data['revenue'].sum()

    # Calculate the total revenue for the hotel in the specified month and year
    total_revenue = channel_mix_data[
        (channel_mix_data['hotel_id'] == hotel_id) & 
        (channel_mix_data['date'].dt.month == month_year.month) & 
        (channel_mix_data['date'].dt.year == month_year.year)
    ]['revenue'].sum()

    # Calculate the percentage of revenue from OTA
    ota_percentage = (ota_revenue / total_revenue * 100) if total_revenue else 0
    return ota_percentage


def calculate_brand_com_conversion(data, hotel_id, month_num, year):
     # Extract visits_and_rev_data DataFrame from the data dictionary
    visits_and_rev_data = data['visits_and_rev_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter data for the specific hotel, month, and year
    filtered_data = visits_and_rev_data[
        (visits_and_rev_data['hotel_id'] == hotel_id) & 
        (visits_and_rev_data['date'].dt.month == month_year.month) & 
        (visits_and_rev_data['date'].dt.year == month_year.year)
    ]

    # Calculate the total traffic and bookings for Brand.com
    brand_com_traffic = filtered_data['traffic'].sum()
    brand_com_bookings = filtered_data['booking'].sum()

    # Calculate Brand.com Conversion Percentage
    if brand_com_traffic > 0:
        brand_com_conversion_percentage = (brand_com_bookings / brand_com_traffic * 100)
    else:
        brand_com_conversion_percentage = 0

    return brand_com_conversion_percentage



def calculate_search_traffic_metrics(data,hotel_id, month_num, year):
     # Extract visits_and_rev_data and source_traffic_data DataFrame from the data dictionary
    visits_and_rev_data = data['visits_and_rev_data']
    source_traffic_data = data['source_traffic_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter for the specific hotel, month, and year in both DataFrames
    filtered_visits_and_rev = visits_and_rev_data[
        (visits_and_rev_data['hotel_id'] == hotel_id) & 
        (visits_and_rev_data['date'].dt.month == month_year.month) & 
        (visits_and_rev_data['date'].dt.year == month_year.year)
    ]
    filtered_source_traffic = source_traffic_data[
        (source_traffic_data['hotel_id'] == hotel_id) & 
        (source_traffic_data['date'].dt.month == month_year.month) & 
        (source_traffic_data['date'].dt.year == month_year.year)
    ]

    # Calculate total traffic and search traffic (organic + paid)
    total_traffic = filtered_visits_and_rev['traffic'].sum()
    organic_search_traffic = filtered_source_traffic[filtered_source_traffic['source_id'].isin(organic_search_ids)]['visits'].sum()
    paid_search_traffic = filtered_source_traffic[filtered_source_traffic['source_id'].isin(paid_search_ids)]['visits'].sum()

    # Calculate the percentage of search traffic
    search_traffic_percentage = ((organic_search_traffic + paid_search_traffic) / total_traffic * 100) if total_traffic else 0

    return search_traffic_percentage

def calculate_metasearch_traffic_percentage(data,hotel_id, month_num, year):
    # Extract visits_and_rev_data and source_traffic_data DataFrame from the data dictionary
    visits_and_rev_data = data['visits_and_rev_data']
    source_traffic_data = data['source_traffic_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter source_traffic_data for the specific hotel and Metasearch IDs
    metasearch_traffic_data = source_traffic_data[
        (source_traffic_data['hotel_id'] == hotel_id) &
        (source_traffic_data['source_id'].isin(metasearch_ids))]

    # Filter for entries in the specified month and year
    metasearch_traffic_data = metasearch_traffic_data[
        (metasearch_traffic_data['date'].dt.month == month_year.month) &
        (metasearch_traffic_data['date'].dt.year == month_year.year)]

    metasearch_traffic = metasearch_traffic_data['visits'].sum()

    # Filter visits_and_rev_data for the specific hotel
    hotel_total_traffic_data = visits_and_rev_data[
        (visits_and_rev_data['hotel_id'] == hotel_id)]

    # Filter for entries in the specified month and year
    hotel_total_traffic_data = hotel_total_traffic_data[
        (hotel_total_traffic_data['date'].dt.month == month_year.month) &
        (hotel_total_traffic_data['date'].dt.year == month_year.year)]

    hotel_total_traffic = hotel_total_traffic_data['traffic'].sum()

    if hotel_total_traffic:
        hotel_metasearch_percentage = (metasearch_traffic / hotel_total_traffic) * 100
    else:
        hotel_metasearch_percentage = 0

    return hotel_metasearch_percentage

def calculate_brand_com_social_percentage(data,hotel_id, month_num, year):
    # Extract visits_and_rev_data and source_traffic_data DataFrame from the data dictionary
    visits_and_rev_data = data['visits_and_rev_data']
    source_traffic_data = data['source_traffic_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter for the specific hotel, month, and year in both DataFrames
    filtered_visits_and_rev = visits_and_rev_data[
        (visits_and_rev_data['hotel_id'] == hotel_id) & 
        (visits_and_rev_data['date'].dt.month == month_year.month) & 
        (visits_and_rev_data['date'].dt.year == month_year.year)
    ]
    filtered_source_traffic = source_traffic_data[
        (source_traffic_data['hotel_id'] == hotel_id) & 
        (source_traffic_data['date'].dt.month == month_year.month) & 
        (source_traffic_data['date'].dt.year == month_year.year)
    ]

    # Calculate total traffic for the specific hotel
    hotel_total_traffic = filtered_visits_and_rev['traffic'].sum()

    # Sum of social traffic (organic + paid)
    social_traffic = filtered_source_traffic[filtered_source_traffic['source_id'].isin(social_traffic_ids)]['visits'].sum()

    # Calculate the Brand.com (Social) percentage
    brand_com_social_percentage = (social_traffic / hotel_total_traffic * 100) if hotel_total_traffic else 0

    return brand_com_social_percentage

def calculate_social_impressions(data,hotel_id, month_num, year):
    # Extract paid_media_data DataFrame from the data dictionary
    paid_media_data = data['paid_media_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter 'Paid Media' data for the specific hotel, month, and year
    hotel_paid_media = paid_media_data[
        (paid_media_data['hotel_id'] == hotel_id) & 
        (paid_media_data['date'].dt.month == month_year.month) & 
        (paid_media_data['date'].dt.year == month_year.year)
    ]

    # Calculate total impressions for Facebook
    facebook_impressions = hotel_paid_media[
        hotel_paid_media['paid_media_source_id'] == 'f9da8533-b300-4ba2-b9b5-e9c7d2d1f5ab']['impression'].sum()

    return facebook_impressions

def calculate_social_spend(data, hotel_id, month_num, year,):
    # Extract paid_media_data DataFrame from the data dictionary
    paid_media_data = data['paid_media_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)

    # Filter 'Paid Media' data for the specific hotel, month, and year
    hotel_social_spend = paid_media_data[
        (paid_media_data['hotel_id'] == hotel_id) & 
        (paid_media_data['date'].dt.month == month_year.month) & 
        (paid_media_data['date'].dt.year == month_year.year)
    ]

    # Calculate total spend for Facebook
    facebook_spend = hotel_social_spend[hotel_social_spend['paid_media_source_id'] == 'f9da8533-b300-4ba2-b9b5-e9c7d2d1f5ab']['spend'].sum()

    return facebook_spend


def calculate_OTA_spend(data, hotel_id, month_num, year):
    # Extract paid_media_data DataFrame from the data dictionary
    paid_media_data = data['paid_media_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)
    # Filter 'Paid Media' data for the specific hotel, month, and year
    hotel_ota_spend = paid_media_data[
        (paid_media_data['hotel_id'] == hotel_id) & 
        (paid_media_data['date'].dt.month == month_year.month) & 
        (paid_media_data['date'].dt.year == month_year.year)
    ]

    # Calculate total spend for OTA
    ota_spend = hotel_ota_spend[hotel_ota_spend['paid_media_source_id'].isin(ota_ids)]['spend'].sum()

    return ota_spend

def calculate_ota_roas(data, hotel_id, month_num, year):
    # Extract paid_media_data DataFrame from the data dictionary
    paid_media_data = data['paid_media_data']
    # Create a datetime object for the first day of the given month and year
    month_year = pd.Timestamp(year=year, month=month_num, day=1)
    # Filter 'Paid Media' data for the specific hotel, month, year, and OTA
    hotel_ota_data = paid_media_data[
        (paid_media_data['hotel_id'] == hotel_id) & 
        (paid_media_data['date'].dt.month == month_year.month) & 
        (paid_media_data['date'].dt.year == month_year.year) & 
        (paid_media_data['paid_media_source_id'].isin(ota_ids))
    ]

    # Calculate total OTA revenue and spend for the specific hotel
    hotel_ota_revenue = hotel_ota_data['revenue'].sum()
    hotel_ota_spend = hotel_ota_data['spend'].sum()

    # Calculate the OTA ROAS
    hotel_ota_roas = hotel_ota_revenue / hotel_ota_spend if hotel_ota_spend else 0

    return hotel_ota_roas

# Calculate and plot metrics for all hotels
months = {
    'Jan': 1,
    'Feb': 2,
    'Mar': 3,
    'Apr': 4,
    'May': 5,
    'Jun': 6,
    'Jul': 7,
    'Aug': 8,
    'Sep': 9,
    'Oct': 10,
    'Nov': 11,
    'Dec': 12
}

metric_functions = [calculate_brand_com_metrics, calculate_ota_metrics, calculate_brand_com_conversion, calculate_search_traffic_metrics, calculate_metasearch_traffic_percentage, calculate_brand_com_social_percentage, calculate_social_impressions,calculate_social_spend,calculate_OTA_spend,calculate_ota_roas]


def plot_histograms_for_metric(metric_values_by_month, metric_name):
    num_months = len(metric_values_by_month)
    fig, axs = plt.subplots(1, num_months, figsize=(20, 5), sharey=True)
    fig.suptitle(f'{metric_name} Histograms (Jan-Aug)')

    # A set of colors to cycle through for each month
    colors = plt.cm.viridis(np.linspace(0, 1, num_months))
    
    for i, (month_name, values) in enumerate(metric_values_by_month.items()):
        axs[i].hist(values, bins=10, color=colors[i], edgecolor='black', alpha=0.7)
        axs[i].set_title(month_name)
        axs[i].set_xlabel('Value')
        axs[i].set_ylabel('Frequency' if i == 0 else '')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.show()

def filter_hotels(data, hotels, metric_functions, year, min_non_zero_metrics=1):
    filtered_hotels = set()
    for hotel_id in hotels:
        non_zero_metrics_count = 0
        for calculate_metric in metric_functions:
            for month_num in range(1, 9):  # January to August
                metric_value = calculate_metric(data, hotel_id, month_num, year)
                if metric_value > 0:
                    non_zero_metrics_count += 1
                    if non_zero_metrics_count >= min_non_zero_metrics:
                        break
            if non_zero_metrics_count >= min_non_zero_metrics:
                break

        if non_zero_metrics_count >= min_non_zero_metrics:
            filtered_hotels.add(hotel_id)

    return filtered_hotels



def calculate_statistics(metric_values):
    metric_series = pd.Series(metric_values)
    return {
        "Sample Size": len(metric_series),
        "Min": metric_series.min(),
        "Max": metric_series.max(),
        "Mean": metric_series.mean(),
        "Median": metric_series.median(),
        "Mode": metric_series.mode()[0] if not metric_series.mode().empty else None,
        "Standard Deviation": metric_series.std(),
        "95% Confidence Interval": stats.norm.interval(0.95, loc=metric_series.mean(), scale=metric_series.std()/np.sqrt(len(metric_series)))
    }

def process_and_plot_data(data, metrics_to_calculate=None, metrics_to_print=None):
    year = 2023
    month_mapping = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug'}
    
    if metrics_to_calculate is None:
        metrics_to_calculate = metric_functions

    if metrics_to_print is None:
        metrics_to_print = {func.__name__ for func in metric_functions}

    unique_hotels = set(data['channel_mix_data']['hotel_id'].unique()) | \
                    set(data['visits_and_rev_data']['hotel_id'].unique()) | \
                    set(data['source_traffic_data']['hotel_id'].unique()) | \
                    set(data['paid_media_data']['hotel_id'].unique())

    for calculate_metric in metrics_to_calculate:
        metric_name = calculate_metric.__name__
        if metric_name in metrics_to_print:
            print(f"\nProcessing metric: {metric_name}")
            metric_values_by_month = {month_mapping[month]: [] for month in range(1, 9)}
            for month in range(1, 9):  # January to August
                filtered_hotels = filter_hotels(data, unique_hotels, [calculate_metric], year)
                metric_values = []
                for hotel_id in filtered_hotels:
                    if calculate_metric == calculate_brand_com_metrics:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_ota_metrics:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_brand_com_conversion:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_search_traffic_metrics:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_metasearch_traffic_percentage:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_brand_com_social_percentage:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_social_impressions:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_social_spend:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_OTA_spend:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    elif calculate_metric == calculate_ota_roas:
                        metric_value = calculate_metric(data, hotel_id, month, year)
                    metric_value = calculate_metric(data,hotel_id,month,year)
                    metric_values_by_month[month_mapping[month]].append(metric_value)
                    metric_values.append(metric_value)
                stats = calculate_statistics(metric_values)
                
               
                print(f"\n{month_mapping[month]} Statistics for {metric_name}:")
                for stat_name, stat_value in stats.items():
                    print(f"{stat_name}: {stat_value}")
                    
            plot_histograms_for_metric(metric_values_by_month, calculate_metric.__name__)

            
def gather_hotel_metrics(data, hotels, metric_functions, year):
    hotel_metrics = {}
    for hotel_id in hotels:
        hotel_data = {}
        for calculate_metric in metric_functions:
            metric_name = calculate_metric.__name__
            metric_data = {}
            for month_num in range(1, 9):  # January to August
                if calculate_metric == calculate_brand_com_metrics:
                    metric_value = calculate_metric(data['channel_mix_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_ota_metrics:
                    metric_value = calculate_metric(data['channel_mix_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_brand_com_conversion:
                    metric_value = calculate_metric(data['visits_and_rev_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_search_traffic_metrics:
                    metric_value = calculate_metric(data['visits_and_rev_data'], data['source_traffic_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_metasearch_traffic_percentage:
                    metric_value = calculate_metric(data['visits_and_rev_data'], data['source_traffic_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_brand_com_social_percentage:
                    metric_value = calculate_metric(data['visits_and_rev_data'], data['source_traffic_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_social_impressions:
                    metric_value = calculate_metric(data['paid_media_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_social_spend:
                    metric_value = calculate_metric(data['paid_media_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_OTA_spend:
                    metric_value = calculate_metric(data['paid_media_data'], hotel_id, month_num, year)
                elif calculate_metric == calculate_ota_roas:
                    metric_value = calculate_metric(data['paid_media_data'], hotel_id, month_num, year)
                metric_data[month_num] = metric_value
            hotel_data[metric_name] = metric_data
        
        if any(metric_value != 0 for metric_data in hotel_data.values() for metric_value in metric_data.values()):
            hotel_metrics[hotel_id] = hotel_data

    return hotel_metrics

# def process_and_save_data():
#     year = 2023
#     channel_mix_data, visits_and_rev_data, source_traffic_data, paid_media_data = load_data()
    
#     # Gather unique hotels from all data sheets
#     unique_hotels = set(channel_mix_data['hotel_id'].unique()) | \
#                     set(visits_and_rev_data['hotel_id'].unique()) | \
#                     set(source_traffic_data['hotel_id'].unique()) | \
#                     set(paid_media_data['hotel_id'].unique())

#     # Gather hotel metrics
#     hotel_metrics = gather_hotel_metrics(unique_hotels, metric_functions, year)

#      # Exclude hotels with zero metrics across all months
#     hotel_metrics = {hotel_id: metrics for hotel_id, metrics in hotel_metrics.items()
#                      if any(value != 0 for month_metrics in metrics.values() for value in month_metrics.values())}

#     # Create DataFrame from hotel metrics
#     df = pd.DataFrame.from_dict({(i,j): hotel_metrics[i][j] 
#                                  for i in hotel_metrics.keys() 
#                                  for j in hotel_metrics[i].keys()},
#                                 orient='index')

#     # Save DataFrame to Excel
#     df.to_excel('hotel_metrics.xlsx')

# # Call the main function
# process_and_save_data()

# Add your metric calculation functions like calculate_brand_com_metrics, calculate_ota_metrics, etc. here

# Call the main function
# Specify which metrics to calculate and print
metrics_to_calculate = [calculate_brand_com_metrics, calculate_ota_metrics, calculate_brand_com_conversion, calculate_search_traffic_metrics, calculate_metasearch_traffic_percentage, calculate_brand_com_social_percentage, calculate_social_impressions,calculate_social_spend,calculate_OTA_spend,calculate_ota_roas]
metrics_to_print = {"calculate_brand_com_metrics", "calculate_ota_metrics", "calculate_brand_com_conversion", "calculate_search_traffic_metrics", "calculate_metasearch_traffic_percentage", "calculate_brand_com_social_percentage", "calculate_social_impressions","calculate_social_spend","calculate_OTA_spend","calculate_ota_roas"}
# Usage

process_and_plot_data(data, metrics_to_calculate, metrics_to_print)








 



