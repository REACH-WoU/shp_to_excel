import os
import zipfile
import geopandas as gpd
import pandas as pd
from shapely.geometry import shape, Point
from pyproj import Transformer
import json
from shiny import App, ui, render

# Function to convert Decimal Degrees (DD) to DMS
def decimal_to_dms(decimal_coord):
    degrees = int(decimal_coord)
    minutes_float = abs(decimal_coord - degrees) * 60
    minutes = int(minutes_float)
    seconds = (minutes_float - minutes) * 60
    return f"{abs(degrees)}\u00b0 {minutes}' {seconds:.2f}\""

def convert_to_esri_json(geometry, tolerance=0.001):
    """
    Convert Shapely Polygon or MultiPolygon geometry to EsriJSON format with 'rings'.
    """
    simplified_geometry = geometry.simplify(tolerance, preserve_topology=True)
    rings = []

    if simplified_geometry.geom_type == "Polygon":
        # Single Polygon: extract exterior coordinates
        coords = list(simplified_geometry.exterior.coords)
        rings.append([list(coord) for coord in coords])
    
    elif simplified_geometry.geom_type == "MultiPolygon":
        # MultiPolygon: iterate through all polygons
        for polygon in simplified_geometry.geoms:
            coords = list(polygon.exterior.coords)
            rings.append([list(coord) for coord in coords])

    return {"rings": rings}

transformer = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)

app_ui = ui.page_fluid(
    ui.h2("Shapefile Processor"),
    ui.input_file("archive", label="Upload archive with shape files", accept=[".zip", ".7z", ".rar"]),
    ui.download_button("download_xlsx", "Download Processed Excel File")
)

def server(input, output, session):

    if not os.path.exists("output_dir"):
        os.mkdir("output_dir")

    @render.download()
    def download_xlsx():
        if os.path.exists("archive_dir"):
            for file in os.listdir("archive_dir"):
                os.remove("archive_dir/" + file)
            os.rmdir("archive_dir")
        
        os.mkdir("archive_dir")


        if input.archive() is not None:
            
            uploaded_file = input.archive()[0]["datapath"]
            
            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                zip_ref.extractall("archive_dir")

            table_data = []

            for root, _, files in os.walk("archive_dir"):
                for file in files:
                    if file.endswith(".shp"):
                        shp_path = os.path.join(root, file)

                        gdf = gpd.read_file(shp_path)

                        for index, row in gdf.iterrows():
                            geometry = row.geometry.centroid
                            lon, lat = geometry.x, geometry.y

                            xDMS = decimal_to_dms(lon)
                            yDMS = decimal_to_dms(lat)

                            xWM, yWM = transformer.transform(lon, lat)

                            esri_json = convert_to_esri_json(row.geometry)
                            compressed_esri_json = json.dumps(esri_json, separators=(",", ":"))
                            compressed_esri_json = compressed_esri_json.replace(' ', '').replace("'", '"')

                            table_data.append({
                                "X": lon,
                                "Y": lat,
                                "xDMS": xDMS,
                                "yDMS": yDMS,
                                "xWM": round(xWM),
                                "yWM": round(yWM),
                                "EsriJSON_Polygons": compressed_esri_json
                            })

            df = pd.DataFrame(table_data)

            output_file = os.path.join("output_dir", "output.xlsx")
            df.to_excel(output_file, index=False)

            return output_file

app = App(app_ui, server)
# app.run()