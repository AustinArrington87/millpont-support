import json
from shapely.geometry import shape, Point
import ijson
import h3
import numpy as np
import pandas as pd
from decimal import Decimal
# Add country name to ISO3 mapping
COUNTRY_TO_ISO3 = {
    # North America
    'United States of America': 'USA',
    'United States': 'USA',
    'USA': 'USA',
    'US': 'USA',
    'Canada': 'CAN',
    'Mexico': 'MEX',
    'Greenland': 'GRL',
    
    # Central America & Caribbean
    'Guatemala': 'GTM',
    'Belize': 'BLZ',
    'Honduras': 'HND',
    'El Salvador': 'SLV',
    'Nicaragua': 'NIC',
    'Costa Rica': 'CRI',
    'Panama': 'PAN',
    'Cuba': 'CUB',
    'Jamaica': 'JAM',
    'Haiti': 'HTI',
    'Dominican Republic': 'DOM',
    'Puerto Rico': 'PRI',
    'Trinidad and Tobago': 'TTO',
    'Barbados': 'BRB',
    'Saint Vincent and the Grenadines': 'VCT',
    'Saint Lucia': 'LCA',
    'Grenada': 'GRD',
    'Dominica': 'DMA',
    'Antigua and Barbuda': 'ATG',
    'Saint Kitts and Nevis': 'KNA',
    'Bahamas': 'BHS',
    
    # South America
    'Brazil': 'BRA',
    'Argentina': 'ARG',
    'Chile': 'CHL',
    'Peru': 'PER',
    'Colombia': 'COL',
    'Venezuela': 'VEN',
    'Ecuador': 'ECU',
    'Bolivia': 'BOL',
    'Paraguay': 'PRY',
    'Uruguay': 'URY',
    'Guyana': 'GUY',
    'Suriname': 'SUR',
    'French Guiana': 'GUF',
    
    # Europe - Western
    'United Kingdom': 'GBR',
    'UK': 'GBR',
    'Great Britain': 'GBR',
    'Britain': 'GBR',
    'England': 'GBR',
    'Scotland': 'GBR',
    'Wales': 'GBR',
    'Northern Ireland': 'GBR',
    'Ireland': 'IRL',
    'France': 'FRA',
    'Germany': 'DEU',
    'Spain': 'ESP',
    'Italy': 'ITA',
    'Portugal': 'PRT',
    'Netherlands': 'NLD',
    'Belgium': 'BEL',
    'Switzerland': 'CHE',
    'Austria': 'AUT',
    'Luxembourg': 'LUX',
    'Monaco': 'MCO',
    'Liechtenstein': 'LIE',
    'San Marino': 'SMR',
    'Vatican City': 'VAT',
    'Andorra': 'AND',
    'Malta': 'MLT',
    'Cyprus': 'CYP',
    
    # Europe - Northern
    'Sweden': 'SWE',
    'Norway': 'NOR',
    'Denmark': 'DNK',
    'Finland': 'FIN',
    'Iceland': 'ISL',
    'Estonia': 'EST',
    'Latvia': 'LVA',
    'Lithuania': 'LTU',
    
    # Europe - Eastern
    'Russia': 'RUS',
    'Russian Federation': 'RUS',
    'Poland': 'POL',
    'Ukraine': 'UKR',
    'Belarus': 'BLR',
    'Czech Republic': 'CZE',
    'Czechia': 'CZE',
    'Slovakia': 'SVK',
    'Hungary': 'HUN',
    'Romania': 'ROU',
    'Bulgaria': 'BGR',
    'Moldova': 'MDA',
    'Slovenia': 'SVN',
    'Croatia': 'HRV',
    'Bosnia and Herzegovina': 'BIH',
    'Serbia': 'SRB',
    'Montenegro': 'MNE',
    'North Macedonia': 'MKD',
    'Macedonia': 'MKD',
    'Albania': 'ALB',
    'Kosovo': 'XKX',
    'Georgia': 'GEO',
    'Armenia': 'ARM',
    'Azerbaijan': 'AZE',
    
    # Asia - East
    'China': 'CHN',
    "People's Republic of China": 'CHN',
    'Japan': 'JPN',
    'South Korea': 'KOR',
    'Korea, South': 'KOR',
    'Republic of Korea': 'KOR',
    'North Korea': 'PRK',
    'Korea, North': 'PRK',
    "Democratic People's Republic of Korea": 'PRK',
    'Mongolia': 'MNG',
    'Taiwan': 'TWN',
    'Hong Kong': 'HKG',
    'Macau': 'MAC',
    
    # Asia - Southeast
    'Indonesia': 'IDN',
    'Philippines': 'PHL',
    'Vietnam': 'VNM',
    'Thailand': 'THA',
    'Myanmar': 'MMR',
    'Burma': 'MMR',
    'Malaysia': 'MYS',
    'Singapore': 'SGP',
    'Cambodia': 'KHM',
    'Laos': 'LAO',
    'Brunei': 'BRN',
    'Timor-Leste': 'TLS',
    'East Timor': 'TLS',
    
    # Asia - South
    'India': 'IND',
    'Pakistan': 'PAK',
    'Bangladesh': 'BGD',
    'Sri Lanka': 'LKA',
    'Nepal': 'NPL',
    'Bhutan': 'BTN',
    'Maldives': 'MDV',
    'Afghanistan': 'AFG',
    
    # Asia - Central
    'Kazakhstan': 'KAZ',
    'Uzbekistan': 'UZB',
    'Turkmenistan': 'TKM',
    'Kyrgyzstan': 'KGZ',
    'Tajikistan': 'TJK',
    
    # Asia - Western (Middle East)
    'Turkey': 'TUR',
    'Iran': 'IRN',
    'Iraq': 'IRQ',
    'Syria': 'SYR',
    'Lebanon': 'LBN',
    'Jordan': 'JOR',
    'Israel': 'ISR',
    'Palestine': 'PSE',
    'Saudi Arabia': 'SAU',
    'Yemen': 'YEM',
    'Oman': 'OMN',
    'United Arab Emirates': 'ARE',
    'UAE': 'ARE',
    'Qatar': 'QAT',
    'Bahrain': 'BHR',
    'Kuwait': 'KWT',
    
    # Africa - North
    'Egypt': 'EGY',
    'Libya': 'LBY',
    'Tunisia': 'TUN',
    'Algeria': 'DZA',
    'Morocco': 'MAR',
    'Sudan': 'SDN',
    'South Sudan': 'SSD',
    
    # Africa - West
    'Nigeria': 'NGA',
    'Ghana': 'GHA',
    'Senegal': 'SEN',
    'Mali': 'MLI',
    'Burkina Faso': 'BFA',
    'Niger': 'NER',
    'Guinea': 'GIN',
    'Sierra Leone': 'SLE',
    'Liberia': 'LBR',
    'Ivory Coast': 'CIV',
    "Côte d'Ivoire": 'CIV',
    'Benin': 'BEN',
    'Togo': 'TGO',
    'Gambia': 'GMB',
    'Guinea-Bissau': 'GNB',
    'Cape Verde': 'CPV',
    'Mauritania': 'MRT',
    
    # Africa - Central
    'Democratic Republic of Congo': 'COD',
    'Dem. Rep. Congo': 'COD',
    'DRC': 'COD',
    'Congo': 'COG',
    'Republic of Congo': 'COG',
    'Central African Republic': 'CAF',
    'Central African Rep.': 'CAF',
    'CAR': 'CAF',
    'Cameroon': 'CMR',
    'Chad': 'TCD',
    'Equatorial Guinea': 'GNQ',
    'Gabon': 'GAB',
    'São Tomé and Príncipe': 'STP',
    
    # Africa - East
    'Ethiopia': 'ETH',
    'Kenya': 'KEN',
    'Tanzania': 'TZA',
    'Uganda': 'UGA',
    'Rwanda': 'RWA',
    'Burundi': 'BDI',
    'Somalia': 'SOM',
    'Djibouti': 'DJI',
    'Eritrea': 'ERI',
    'Madagascar': 'MDG',
    'Mauritius': 'MUS',
    'Seychelles': 'SYC',
    'Comoros': 'COM',
    'Malawi': 'MWI',
    'Mozambique': 'MOZ',
    'Zambia': 'ZMB',
    'Zimbabwe': 'ZWE',
    
    # Africa - Southern
    'South Africa': 'ZAF',
    'Botswana': 'BWA',
    'Namibia': 'NAM',
    'Lesotho': 'LSO',
    'Eswatini': 'SWZ',
    'Swaziland': 'SWZ',
    'Angola': 'AGO',
    
    # Oceania
    'Australia': 'AUS',
    'New Zealand': 'NZL',
    'Papua New Guinea': 'PNG',
    'Fiji': 'FJI',
    'Solomon Islands': 'SLB',
    'Vanuatu': 'VUT',
    'Samoa': 'WSM',
    'Tonga': 'TON',
    'Kiribati': 'KIR',
    'Tuvalu': 'TUV',
    'Nauru': 'NRU',
    'Palau': 'PLW',
    'Marshall Islands': 'MHL',
    'Micronesia': 'FSM',
    'Federated States of Micronesia': 'FSM',
    
    # Common alternative names and abbreviations
    'United States of America': 'USA',
    'Republic of Korea': 'KOR',
    'Democratic Republic of the Congo': 'COD',
    'Republic of the Congo': 'COG',
    'Czech Republic': 'CZE',
    'Slovak Republic': 'SVK',
    'Republic of China': 'TWN',  # Taiwan
    'Republic of South Africa': 'ZAF',
    'Kingdom of Saudi Arabia': 'SAU',
    'Islamic Republic of Iran': 'IRN',
    'Republic of India': 'IND',
    'Federative Republic of Brazil': 'BRA',
    'Russian Federation': 'RUS',
    'Federal Republic of Germany': 'DEU',
    'French Republic': 'FRA',
    'Italian Republic': 'ITA',
    'Kingdom of Spain': 'ESP',
    'Portuguese Republic': 'PRT',
    'Kingdom of the Netherlands': 'NLD',
    'Kingdom of Belgium': 'BEL',
    'Swiss Confederation': 'CHE',
    'Republic of Austria': 'AUT',
    'United Kingdom of Great Britain and Northern Ireland': 'GBR',
    'Republic of Ireland': 'IRL',
    'Kingdom of Sweden': 'SWE',
    'Kingdom of Norway': 'NOR',
    'Kingdom of Denmark': 'DNK',
    'Republic of Finland': 'FIN',
    'Republic of Iceland': 'ISL',
    'Commonwealth of Australia': 'AUS',
    'New Zealand': 'NZL',
    'Dominion of New Zealand': 'NZL',
    'Independent State of Papua New Guinea': 'PNG',
    'Republic of the Philippines': 'PHL',
    'Socialist Republic of Vietnam': 'VNM',
    'Kingdom of Thailand': 'THA',
    'Republic of Indonesia': 'IDN',
    'Malaysia': 'MYS',
    'Republic of Singapore': 'SGP',
    'Kingdom of Cambodia': 'KHM',
    'Lao People\'s Democratic Republic': 'LAO',
    'Union of Myanmar': 'MMR',
    'People\'s Republic of Bangladesh': 'BGD',
    'Democratic Socialist Republic of Sri Lanka': 'LKA',
    'Federal Democratic Republic of Nepal': 'NPL',
    'Kingdom of Bhutan': 'BTN',
    'Republic of the Maldives': 'MDV',
    'Islamic Republic of Afghanistan': 'AFG',
    'Republic of Kazakhstan': 'KAZ',
    'Republic of Uzbekistan': 'UZB',
    'Republic of Turkmenistan': 'TKM',
    'Kyrgyz Republic': 'KGZ',
    'Republic of Tajikistan': 'TJK',
    'Republic of Turkey': 'TUR',
    'Islamic Republic of Iran': 'IRN',
    'Republic of Iraq': 'IRQ',
    'Syrian Arab Republic': 'SYR',
    'Lebanese Republic': 'LBN',
    'Hashemite Kingdom of Jordan': 'JOR',
    'State of Israel': 'ISR',
    'State of Palestine': 'PSE',
    'Kingdom of Saudi Arabia': 'SAU',
    'Republic of Yemen': 'YEM',
    'Sultanate of Oman': 'OMN',
    'United Arab Emirates': 'ARE',
    'State of Qatar': 'QAT',
    'Kingdom of Bahrain': 'BHR',
    'State of Kuwait': 'KWT',
    'Arab Republic of Egypt': 'EGY',
    'State of Libya': 'LBY',
    'Republic of Tunisia': 'TUN',
    'People\'s Democratic Republic of Algeria': 'DZA',
    'Kingdom of Morocco': 'MAR',
    'Republic of the Sudan': 'SDN',
    'Republic of South Sudan': 'SSD'
}

H3_RESOLUTION = 5  # ~8km hexes

def to_jsonable(x):
    if x is None or isinstance(x, (bool, int, float, str)):
        return x
    if isinstance(x, Decimal):
        try:
            if x == x.to_integral_value():
                return int(x)
        except Exception:
            pass
        return float(x)
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (np.floating,)):
        return float(x)
    if isinstance(x, dict):
        return {k: to_jsonable(v) for k, v in x.items()}
    if isinstance(x, (list, tuple, set)):
        return [to_jsonable(v) for v in x]
    return str(x)

def get_h3_indices(geometry):
    """Get approximate H3 cells covering geometry."""
    minx, miny, maxx, maxy = geometry.bounds
    indices = set()
    lat_step = (maxy - miny) / 10 if maxy > miny else 1
    lng_step = (maxx - minx) / 10 if maxx > minx else 1
    try:
        polygons = list(geometry.geoms) if geometry.geom_type == 'MultiPolygon' else [geometry]
        for polygon in polygons:
            for lon, lat in polygon.exterior.coords:
                indices.add(h3.latlng_to_cell(lat, lon, H3_RESOLUTION))
            for lat in np.arange(miny, maxy, lat_step):
                for lon in np.arange(minx, maxx, lng_step):
                    if polygon.contains(Point(lon, lat)):
                        indices.add(h3.latlng_to_cell(lat, lon, H3_RESOLUTION))
    except Exception as e:
        print(f"[H3] Error for {geometry.geom_type}: {e}")
    return indices

def get_target_countries_from_geojson(geojson_file):
    print("Extracting countries from GeoJSON...")
    countries = set()
    with open(geojson_file, 'rb') as f:
        for feature in ijson.items(f, 'features.item'):
            country = feature.get('properties', {}).get('country')
            if country:
                countries.add(country)
    unique_countries = list(countries)
    iso3_codes = [COUNTRY_TO_ISO3[c] for c in unique_countries if c in COUNTRY_TO_ISO3]
    print(f"Found countries: {unique_countries}")
    print(f"ISO3 codes: {iso3_codes}")
    return iso3_codes

def load_protected_areas(geojson_file, target_countries=None):
    print("Loading protected areas...")
    protected_areas_index = {}
    total, kept = 0, 0
    with open(geojson_file, 'rb') as f:
        for feature in ijson.items(f, 'features.item'):
            total += 1
            props = feature.get('properties', {})
            if target_countries and props.get('ISO3') not in target_countries:
                continue
            geom_dict = feature.get('geometry')
            if not geom_dict:
                continue
            try:
                geom = shape(geom_dict)
            except Exception:
                continue
            area = {
                'WDPAID': to_jsonable(props.get('WDPAID')),
                'NAME': to_jsonable(props.get('NAME')),
                'DESIG_ENG': to_jsonable(props.get('DESIG_ENG')),
                'IUCN_CAT': to_jsonable(props.get('IUCN_CAT')),
                'MARINE': to_jsonable(props.get('MARINE')),
                'STATUS': to_jsonable(props.get('STATUS')),
                'STATUS_YR': to_jsonable(props.get('STATUS_YR')),
                'ISO3': to_jsonable(props.get('ISO3')),
                'geometry': geom,
            }
            for idx in get_h3_indices(geom):
                protected_areas_index.setdefault(idx, []).append(area)
            kept += 1
            if total % 1000 == 0:
                print(f"Processed {total} areas, kept {kept}")
    print(f"Loaded {kept}/{total} protected areas")
    return protected_areas_index

def process_projects_to_csv(projects_geojson_file, protected_areas_index, output_csv):
    print("Processing projects to CSV...")
    rows = []
    processed, overlaps, errors = 0, 0, 0

    with open(projects_geojson_file, 'rb') as f:
        for idx, feature in enumerate(ijson.items(f, 'features.item')):
            try:
                geom = shape(feature['geometry'])
                props = feature.get('properties', {})
                project_id = props.get('id') or feature.get('id')  # support either location

                project_h3 = get_h3_indices(geom)
                pa = {
                    'id': project_id,
                    'PA_WDPAID': None,
                    'PA_NAME': None,
                    'PA_DESIG_ENG': None,
                    'PA_IUCN_CAT': None,
                    'PA_MARINE': None,
                    'PA_STATUS': None,
                    'PA_STATUS_YR': None,
                    'PA_ISO3': None,
                    'unep_overlap': False,
                }

                checked = set()
                found = False
                for cell in project_h3:
                    if found:
                        break
                    for area in protected_areas_index.get(cell, []):
                        wdpaid = area['WDPAID']
                        if wdpaid in checked:
                            continue
                        checked.add(wdpaid)
                        if geom.intersects(area['geometry']):
                            pa.update({
                                'PA_WDPAID': to_jsonable(area['WDPAID']),
                                'PA_NAME': to_jsonable(area['NAME']),
                                'PA_DESIG_ENG': to_jsonable(area['DESIG_ENG']),
                                'PA_IUCN_CAT': to_jsonable(area['IUCN_CAT']),
                                'PA_MARINE': to_jsonable(area['MARINE']),
                                'PA_STATUS': to_jsonable(area['STATUS']),
                                'PA_STATUS_YR': to_jsonable(area['STATUS_YR']),
                                'PA_ISO3': to_jsonable(area['ISO3']),
                                'unep_overlap': True,
                            })
                            found = True
                            break

                rows.append(pa)
                processed += 1
                if pa['unep_overlap']:
                    overlaps += 1
                if processed % 100 == 0:
                    print(f"Processed {processed} projects...")
            except Exception as e:
                errors += 1
                if errors < 20:
                    print(f"Error on feature {idx}: {e}")

    df = pd.DataFrame(rows, columns=[
        'id',
        'PA_WDPAID',
        'PA_NAME',
        'PA_DESIG_ENG',
        'PA_IUCN_CAT',
        'PA_MARINE',
        'PA_STATUS',
        'PA_STATUS_YR',
        'PA_ISO3',
        'unep_overlap',
    ])
    df.to_csv(output_csv, index=False)
    print(f"\nProcessed: {processed}")
    print(f"Overlaps found: {overlaps}")
    print(f"Errors: {errors}")
    print(f"CSV saved to: {output_csv}")

def main():
    protected_areas_file = "geojson/WDPA_Mar2025_Public_merged_polygons.geojson"
    projects_file = "sources_20251022_195516.geojson"
    output_csv = "projects_with_protected_areas.csv"

    target_countries = get_target_countries_from_geojson(projects_file)
    print(f"Filtering protected areas for: {target_countries}")

    pa_index = load_protected_areas(protected_areas_file, target_countries)
    process_projects_to_csv(projects_file, pa_index, output_csv)

if __name__ == "__main__":
    main()
