# Scraping y Análisis de Precios

## Descripción

Este proyecto consiste en un **script de Python** que permite extraer precios de productos desde **MercadoLibre Colombia**, calcular promedios y generar un archivo **Excel** con formato profesional.  
El objetivo principal es obtener información rápida y estructurada sobre precios de productos específicos para análisis, comparación o toma de decisiones.

---

## Funcionalidades

- Scraping de precios de productos en MercadoLibre Colombia.  
- Cálculo de:  
  - Promedio de precios  
  - Promedio con 20% de descuento  
  - Promedio con 30% de descuento  
- Exportación de resultados a un archivo **Excel** con:  
  - Encabezados destacados  
  - Formato de miles (COP)  
  - Ajuste automático de ancho de columnas  
- Código modular, limpio y fácil de mantener.

---

## Requisitos

- Python 3.8 o superior  
- Librerías necesarias:  
```bash
pip install requests openpyxl
