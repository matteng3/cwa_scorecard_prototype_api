from flask import Flask, request, jsonify
import pandas as pd

app = Flask(__name__)

# Load Excel data
excel_path = 'Sample data scorecard (1).xlsx'
channel_mix_data = pd.read_excel(excel_path, sheet_name='ChannelMix')
visits_and_rev_data = pd.read_excel(excel_path, sheet_name='Visits and Rev')
source_traffic_data = pd.read_excel(excel_path, sheet_name='Source Traffic')
paid_media_data = pd.read_excel(excel_path, sheet_name='Paid Media')
google_ad_performance_data = pd.read_excel(excel_path, sheet_name='Google Ads Performance')

# Complete lists for Organic and Paid Search Traffic without "(lowercase)"
organic_search_categories = [
    'NATURAL SEARCH', 'ORGANIC META', 'UNPAID REFERRER', 'SOCIAL MEDIA', 'NON PAID LISTING',
    'NATURAL SOCIAL', 'HOTEL SITES & MICROSITES', 'ORGANIC METASEARCH', 'ORGANIC SOCIAL',
    'Natural Search', 'Natural Social', 'ORGANIC SEARCH', 'UNPAID REFERRERS', 'DIRECTORY LISTINGS',
    'INDEPENDENT WEBSITES', 'LOCAL SEARCH', 'TYPED/BOOKMARKED', 'hotel wifi', 'Referring Domains',
    'Hotel Sites & Microsites', 'ECONFO AND PRE-ARRIVAL EMAIL', 'REFERRED TRAFFIC', 'Partnerships',
    'FRIENDLY URLS', 'No Marketing', 'Unspecified', 'None', 'Email', 'Other Links', 'Natural Social',
    'Display', 'Organic Meta', 'Organic Social', 'email', 'misc', 'unpaid referrers', 'typed/bookmarked', 'hotel wifi',
    'natural search', 'Friendly URLs', 'Referring Domains', '1Hotel Sites & Microsites', 'DIRECTORY LISTINGS',
    'INDEPENDENT WEBSITES', 'LOCAL SEARCH', 'REFERRED TRAFFIC', 'OTHER', 'Other'
]

paid_search_categories = [
    'PAID SEARCH', 'PAID META', 'PAID SOCIAL', 'PAID LISTING', 'DISPLAY', 'AFFILIATES', 'PAID METASEARCH',
    'Paid Search', 'Paid Social', 'Display', 'paid meta', 'paid social', 'paid listing', 
    'SECONDARY AND MINI SITES', 'PARTNERSHIP', 'META SEARCH', 'FRIENDLY URLS', 'meta search', 'RESLINK', 'XXXAAA',
    'E-MAIL OR OTHER OWNED CHANNELS', 'PARTNERSHIPS', 'Paid Meta', 'Paid Social', 'Paid Listing', 'paid search', 
    'Display', 'paid metasearch', 'SECONDARY AND MINI SITES', 'PARTNERSHIP', 'META SEARCH'
]

# Define the organic and paid metasearch categories based on the image provided.
organic_metasearch_categories = [
    'SOCIAL MEDIA', 'NATURAL SOCIAL', 'ORGANIC SOCIAL', 'social media', 'Natural Social','Organic Meta','ORGANIC META','ORGANIC METASEARCH'
]

paid_metasearch_categories = [
    'PAID SOCIAL', 'paid social', 'PAID META','paid metasearch','Paid Social','Paid Meta', 'PAID METASEARCH', 'META SEARCH',
]

social_traffic_categories = [
    'SOCIAL MEDIA', 'PAID SOCIAL', 'NATURAL SOCIAL', 'ORGANIC SOCIAL', 
    'social media', 'Paid Social', 'Natural Social', 'Organic Social',
    'paid social', 'organic social', 'Natural social', 'Paid social'
]

@app.route('/calculate_metrics', methods=['POST'])
def calculate_metrics():
    data = request.json
    selected_hotel = data['selected_hotel']
    selected_month = data['selected_month']
    selected_year = data['selected_year']

    # Perform the calculations for all metrics
    brand_com_percentage, brand_com_score = calculate_brand_com_metrics(selected_hotel, selected_month, selected_year)
    ota_percentage, ota_score = calculate_ota_metrics(selected_hotel, selected_month, selected_year)
    brand_com_conversion_percentage, brand_com_conversion_score = calculate_brand_com_conversion(selected_hotel, selected_month, selected_year)
    search_traffic_percentage, search_traffic_score = calculate_search_traffic_metrics(selected_hotel, selected_month, selected_year)
    hotel_metasearch_percentage, hotel_metasearch_score = calculate_metasearch_traffic_percentage(selected_hotel, selected_month, selected_year)
    brand_com_social_percentage, brand_com_social_score = calculate_brand_com_social_percentage(selected_hotel, selected_month, selected_year)
    facebook_impressions, facebook_impressions_score = calculate_social_impressions(selected_hotel, selected_month, selected_year)
    facebook_spend, facebook_spend_score = calculate_social_spend(selected_hotel, selected_month, selected_year)
    ota_spend, ota_spend_score = calculate_OTA_spend(selected_hotel, selected_month, selected_year)
    hotel_ota_roas, average_ota_roas = calculate_ota_roas(selected_hotel, selected_month, selected_year)
    


    return jsonify({
        'brand_com_percentage': brand_com_percentage,
        'brand_com_score': brand_com_score,
        'ota_percentage': ota_percentage,
        'ota_score': ota_score,
        'brand_com_conversion_percentage': brand_com_conversion_percentage,
        'brand_com_conversion_score': brand_com_conversion_score,
        'search_traffic_percentage': search_traffic_percentage,
        'search_traffic_score': search_traffic_score,
        'hotel_metasearch_percentage': hotel_metasearch_percentage,
        'hotel_metasearch_score': hotel_metasearch_score,
        'brand_com_social_percentage': brand_com_social_percentage,
        'brand_com_social_score': brand_com_social_score,
        'facebook_impressions': facebook_impressions,
        'facebook_impressions_score': facebook_impressions_score,
        'facebook_spend': facebook_spend,
        'facebook_spend_score': facebook_spend_score,
        'ota_spend': ota_spend,
        'ota_spend_score': ota_spend_score,
        'hotel_ota_roas': hotel_ota_roas,
        'average_ota_roas': average_ota_roas
    })

# Function definitions for calculations

def calculate_brand_com_metrics(hotel_name, month, year):
    print("Working on brand.com  %")

    # Convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in both DataFrames to lowercase
    channel_mix_data['s_Month'] = channel_mix_data['s_Month'].str.lower()

    # Convert 's_ChannelType' to lowercase for case-insensitive comparison
    channel_mix_data['s_ChannelType'] = channel_mix_data['s_ChannelType'].str.lower()

    # Logic for calculating Brand.com % of Revenue
    brand_com_revenue = channel_mix_data[
        (channel_mix_data['s_Code'] == hotel_name) & 
        (channel_mix_data['s_Month'] == month) & 
        (channel_mix_data['s_Year'] == year) & 
        (channel_mix_data['s_ChannelType'] == 'web')  # 'web' in lowercase
    ]['d_Revenue'].sum()

    total_revenue = channel_mix_data[
        (channel_mix_data['s_Code'] == hotel_name) & 
        (channel_mix_data['s_Month'] == month) & 
        (channel_mix_data['s_Year'] == year)
    ]['d_Revenue'].sum()

    # Print the values and the formula
    print(f"Brand.com Revenue for {hotel_name} in {month} {year}: {brand_com_revenue}")
    print(f"Total Revenue for {hotel_name} in {month} {year}: {total_revenue}")

    # Calculate the Brand.com Percentage
    if total_revenue:
        brand_com_percentage = (brand_com_revenue / total_revenue * 100)
        print(f"Brand.com Percentage: ({brand_com_revenue} / {total_revenue}) * 100 = {brand_com_percentage}%")
    else:
        brand_com_percentage = 0
        print("Brand.com Percentage: No total revenue to calculate percentage.")

    brand_com_score = simplistic_scoring(brand_com_percentage)
    return brand_com_percentage, brand_com_score

def calculate_ota_metrics(hotel_name, month, year):

    print("Working on OTA metric %")

     # Convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in both DataFrames to lowercase
    channel_mix_data['s_Month'] = channel_mix_data['s_Month'].str.lower()
    channel_mix_data['s_ChannelType'] = channel_mix_data['s_ChannelType'].str.lower()


    # Logic for calculating OTA % of Revenue
    ota_revenue = channel_mix_data[(channel_mix_data['s_Code'] == hotel_name) & 
                                   (channel_mix_data['s_Month'] == month) & 
                                   (channel_mix_data['s_Year'] == year) & 
                                   (channel_mix_data['s_ChannelType'] == 'ota')]['d_Revenue'].sum()
    
    print("Total OTA:", ota_revenue)

    total_revenue = channel_mix_data[(channel_mix_data['s_Code'] == hotel_name) & 
                                    (channel_mix_data['s_Month'] == month) & 
                                    (channel_mix_data['s_Year'] == year)]['d_Revenue'].sum()
    
     # Print the values and the formula
    print(f"OTA Revenue for {hotel_name} in {month} {year}: {ota_revenue}")
    print(f"Total Revenue for {hotel_name} in {month} {year}: {total_revenue}")
    
    if total_revenue:
        ota_percentage = (ota_revenue / total_revenue * 100)
        print(f"OTA Percentage: ({ota_revenue} / {total_revenue}) * 100 = {ota_percentage}%")
    else:
        ota_percentage = 0
        print("OTA Percentage: No total revenue to calculate percentage.")
    ota_score = simplistic_scoring(ota_percentage)
    return ota_percentage, ota_score

def calculate_brand_com_conversion(hotel_name, month, year):

    print("Working on brand.com conversion %")

    # Convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in both DataFrames to lowercase
    visits_and_rev_data['s_Month'] = visits_and_rev_data['s_Month'].str.lower()

    # Logic for calculating Brand.com Conversion %
    brand_com_traffic = visits_and_rev_data[(visits_and_rev_data['s_Code'] == hotel_name) & 
                                            (visits_and_rev_data['s_Month'] == month) & 
                                            (visits_and_rev_data['s_Year'] == year)]['d_Traffic'].sum()
    brand_com_bookings = visits_and_rev_data[(visits_and_rev_data['s_Code'] == hotel_name) & 
                                             (visits_and_rev_data['s_Month'] == month) & 
                                             (visits_and_rev_data['s_Year'] == year)]['d_Booking'].sum()
    
    # Print the values and the formula
    print(f"Brand Conversion for {hotel_name} in {month} {year}: {brand_com_traffic}")
    print(f"Brand Bookings for {hotel_name} in {month} {year}: {brand_com_bookings}")

    if brand_com_bookings:
        brand_com_conversion_percentage = (brand_com_bookings / brand_com_traffic * 100)
        print(f"Brand.com conversion percentage: ({brand_com_bookings} / {brand_com_traffic}) * 100 = {brand_com_conversion_percentage}%")
    else:
        brand_com_conversion_percentage = 0
        print("Brand.com conversion percentage: No total revenue to calculate percentage.")
    brand_com_conversion_score = simplistic_scoring(brand_com_conversion_percentage)
    return brand_com_conversion_percentage, brand_com_conversion_score

def calculate_search_traffic_metrics(hotel_name, month, year):

    print("Working on search traffic%")

    # Convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in both DataFrames to lowercase
    source_traffic_data['s_month'] = source_traffic_data['s_month'].str.lower()
    visits_and_rev_data['s_Month'] = visits_and_rev_data['s_Month'].str.lower()

    # Logic for calculating % of Search Traffic (Organic + Paid)
    hotel_data_source = source_traffic_data[(source_traffic_data['s_codes'] == hotel_name) & 
                                            (source_traffic_data['s_month'] == month) & 
                                            (source_traffic_data['s_year'] == year)]
    total_traffic = visits_and_rev_data[(visits_and_rev_data['s_Code'] == hotel_name) & 
                                        (visits_and_rev_data['s_Month'] == month) & 
                                        (visits_and_rev_data['s_Year'] == year)]['d_Traffic'].sum()
    organic_search_traffic = hotel_data_source[hotel_data_source['s_source'].isin(organic_search_categories)]['d_visits'].sum()
    paid_search_traffic = hotel_data_source[hotel_data_source['s_source'].isin(paid_search_categories)]['d_visits'].sum()

    # Print the values and the formula
    print(f"Search Traffic data for {hotel_name} in {month} {year}: {organic_search_traffic} + {paid_search_traffic}")
    print(f"Total traffic for {hotel_name} in {month} {year}: {total_traffic}")   

    if total_traffic:
        search_traffic_percentage = ((organic_search_traffic + paid_search_traffic) / total_traffic * 100)
        print(f"Search traffic percentage: ({organic_search_traffic} + {paid_search_traffic} / {total_traffic}) * 100 = {search_traffic_percentage}%")
    else:
        search_traffic_percentage = 0
        print("Search traffic percentage: No total revenue to calculate percentage.")
    search_traffic_score = simplistic_scoring(search_traffic_percentage)
    return search_traffic_percentage, search_traffic_score


def calculate_metasearch_traffic_percentage(hotel_name, month, year):
# Logic for calculating % of Metasearch Traffic (Organic + Paid)
    print("Working on metasearch %")

    # Convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in both DataFrames to lowercase
    source_traffic_data['s_month'] = source_traffic_data['s_month'].str.lower()
    visits_and_rev_data['s_Month'] = visits_and_rev_data['s_Month'].str.lower()
    
    # Filter the source traffic for the specific hotel, month, and year
    hotel_source_traffic = source_traffic_data[
        (source_traffic_data['s_codes'] == hotel_name) & 
        (source_traffic_data['s_month'] == month) & 
        (source_traffic_data['s_year'] == year)]
    
    # Calculate total traffic for the specific hotel from the 'Visits and Rev' sheet
    hotel_total_traffic = visits_and_rev_data[
        (visits_and_rev_data['s_Code'] == hotel_name) & 
        (visits_and_rev_data['s_Month'] == month) & 
        (visits_and_rev_data['s_Year'] == year)
    ]['d_Traffic'].sum()

    organic_metasearch_traffic = hotel_source_traffic[hotel_source_traffic['s_source'].isin(organic_metasearch_categories)]['d_visits'].sum()
    paid_metasearch_traffic = hotel_source_traffic[hotel_source_traffic['s_source'].isin(paid_metasearch_categories)]['d_visits'].sum()

    # Print the values and the formula
    print(f"Metasearch Traffic for {hotel_name} in {month} {year}: {organic_metasearch_traffic} + {paid_metasearch_traffic}")
    print(f"Total Traffic for {hotel_name} in {month} {year}: {hotel_total_traffic}")   

    if hotel_total_traffic:
        hotel_metasearch_percentage = ((organic_metasearch_traffic + paid_metasearch_traffic) / hotel_total_traffic * 100)
        print(f"Metasearch Traffic Percentage: ({organic_metasearch_traffic} + {paid_metasearch_traffic} / {hotel_total_traffic}) * 100 = {hotel_metasearch_percentage}%")
    else:
        hotel_metasearch_percentage = 0
        print("Metasearch Traffic Percentage: No total traffic to calculate percentage.")

    hotel_metasearch_score = simplistic_scoring(hotel_metasearch_percentage)
    return hotel_metasearch_percentage, hotel_metasearch_score

def calculate_brand_com_social_percentage(hotel_name, month, year):
    print("Working on Brand.com (Social) metric")

    # Convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in both DataFrames to lowercase
    source_traffic_data['s_month'] = source_traffic_data['s_month'].str.lower()
    visits_and_rev_data['s_Month'] = visits_and_rev_data['s_Month'].str.lower()

    # Filter the source traffic for the specific hotel, month, and year
    hotel_source_traffic = source_traffic_data[
        (source_traffic_data['s_codes'] == hotel_name) & 
        (source_traffic_data['s_month'] == month) & 
        (source_traffic_data['s_year'] == year)]

    # Calculate total traffic for the specific hotel from the 'Visits and Rev' sheet
    hotel_total_traffic = visits_and_rev_data[
        (visits_and_rev_data['s_Code'] == hotel_name) & 
        (visits_and_rev_data['s_Month'] == month) & 
        (visits_and_rev_data['s_Year'] == year)
    ]['d_Traffic'].sum()

    # Sum of social traffic (organic + paid)
    social_traffic = hotel_source_traffic[hotel_source_traffic['s_source'].isin(social_traffic_categories)]['d_visits'].sum()

    # Print the values and the formula
    print(f"Social Traffic for {hotel_name} in {month} {year}: {social_traffic}")
    print(f"Total Traffic for {hotel_name} in {month} {year}: {hotel_total_traffic}")

    # Calculate the Brand.com (Social) percentage
    if hotel_total_traffic:
        brand_com_social_percentage = (social_traffic / hotel_total_traffic * 100)
        print(f"Brand.com (Social) Traffic Percentage: ({social_traffic} / {hotel_total_traffic}) * 100 = {brand_com_social_percentage}%")
    else:
        brand_com_social_percentage = 0
        print("Brand.com (Social) Traffic Percentage: No total traffic to calculate percentage.")

    brand_com_social_score = simplistic_scoring(brand_com_social_percentage)
    return brand_com_social_percentage, brand_com_social_score

def calculate_social_impressions(hotel_name, month, year):
    print("Working on Social Impressions metric")

    # Also convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in the DataFrame to lowercase
    paid_media_data['s_Month'] = paid_media_data['s_Month'].str.lower()

    # Filter 'Paid Media' data for the specific hotel, month, and year
    hotel_paid_media = paid_media_data[
        (paid_media_data['s_Hotelcode'] == hotel_name) & 
        (paid_media_data['s_Month'] == month) & 
        (paid_media_data['s_Year'] == year)
    ]

    print("We've found:", hotel_paid_media)
    # Calculate total impressions for Facebook and Instagram
    facebook_impressions = hotel_paid_media[
        hotel_paid_media['s_MediaChannel'].isin(['Facebook'])
    ]['n_Impression'].sum()

    print("Total impressions:", facebook_impressions)

    # Print the values
    print(f"Facebook Impressions for {hotel_name} in {month} {year}: {facebook_impressions}")

    facebook_impressions_score = simplistic_scoring(facebook_impressions)
    return facebook_impressions, facebook_impressions_score

def calculate_social_spend(hotel_name, month, year):
    print("Working on Social Spend metric")

     # Also convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in the DataFrame to lowercase
    paid_media_data['s_Month'] = paid_media_data['s_Month'].str.lower()

    # Filter 'Paid Media' data for the specific hotel, month, and year
    hotel_social_spend = paid_media_data[
        (paid_media_data['s_Hotelcode'] == hotel_name) & 
        (paid_media_data['s_Month'] == month) & 
        (paid_media_data['s_Year'] == year)
    ]

    print("We've found:", hotel_social_spend)
    # Calculate total impressions for Facebook and Instagram
    facebook_spend = hotel_social_spend[
        hotel_social_spend['s_MediaChannel'].isin(['Facebook'])
    ]['d_Spend'].sum()

    print("Total Spend:", facebook_spend)

    # Print the values
    print(f"Facebook Spend for {hotel_name} in {month} {year}: {facebook_spend}")

    facebook_spend_score = simplistic_scoring(facebook_spend)
    return facebook_spend, facebook_spend_score


def calculate_OTA_spend(hotel_name, month, year):
    print("Working on OTA Spend metric")

    # Also convert the month parameter to lowercase for matching
    month = month.lower()

    # Standardize the month values in the DataFrame to lowercase
    paid_media_data['s_Month'] = paid_media_data['s_Month'].str.lower()

    # Filter 'Paid Media' data for the specific hotel, month, and year
    hotel_ota_spend = paid_media_data[
        (paid_media_data['s_Hotelcode'] == hotel_name) & 
        (paid_media_data['s_Month'] == month) & 
        (paid_media_data['s_Year'] == year)
    ]

    print("We've found:", hotel_ota_spend)
    # Calculate total impressions for OTA
    ota_spend = hotel_ota_spend[
        hotel_ota_spend['s_PaidMedia'] == 'OTA']['d_Spend'].sum()

    print("Total OTA Spend:", ota_spend)

    # Print the values
    print(f"OTA Spend for {hotel_name} in {month} {year}: {ota_spend}")

    ota_spend_score = simplistic_scoring(ota_spend)
    return ota_spend, ota_spend_score

def calculate_ota_roas(hotel_name, month, year):
    print("Working on OTA ROAS metric")

    # Standardize the month values in the DataFrame to lowercase for consistency
    paid_media_data['s_Month'] = paid_media_data['s_Month'].str.lower()
    month = month.lower()

    # Filter 'Paid Media' data for OTA
    ota_data = paid_media_data[
        (paid_media_data['s_Month'] == month) & 
        (paid_media_data['s_Year'] == year) & 
        (paid_media_data['s_PaidMedia'] == 'OTA')
    ]

    # Calculate total OTA revenue and spend for the specific hotel
    hotel_ota_data = ota_data[ota_data['s_Hotelcode'] == hotel_name]
    hotel_ota_revenue = hotel_ota_data['d_Revenue'].sum()
    hotel_ota_spend = hotel_ota_data['d_Spend'].sum()
    hotel_ota_roas = hotel_ota_revenue / hotel_ota_spend if hotel_ota_spend else 0

    # Calculate total OTA revenue and spend for all hotels
    total_ota_revenue = ota_data['d_Revenue'].sum()
    total_ota_spend = ota_data['d_Spend'].sum()
    average_ota_roas = total_ota_revenue / total_ota_spend if total_ota_spend else 0

    print(f"OTA ROAS for {hotel_name} in {month} {year}: {hotel_ota_roas}")
    print(f"Average OTA ROAS for all hotels in {month} {year}: {average_ota_roas}")

    return hotel_ota_roas, average_ota_roas

# Load the Excel file data into pandas dataframes
# source_traffic_data = pd.read_excel('path_to_source_traffic_file.xlsx', sheet_name='Source Traffic')
# visits_and_rev_data = pd.read_excel('path_to_visits_and_rev_file.xlsx', sheet_name='Visits and Rev')

# Example usage (you would replace 'Hotel Name', 'Month', 'Year' with actual values)
# hotel_metasearch_percentage, average_metasearch_percentage = calculate_metasearch_traffic_percentage(
#     'Hotel Name', 'Month', 'Year', source_traffic_data, visits_and_rev_data
# )
# print(hotel_metasearch_percentage, average_metasearch_percentage)

# Example simplistic scoring function; replace with your actual logic
def simplistic_scoring(value):
    return min(max(int(value / 10), 1), 10)

if __name__ == '__main__':
    app.run(debug=True)


