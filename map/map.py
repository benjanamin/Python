import folium

#create map object
m = folium.Map(location=[-33.43870245209876, -70.59427549440213], zoom_start=20)

#Generar mapa
m.save('map.html')