# Copaco Data Verrijking - Excel Export Tool

Een webgebaseerde tool voor het verwerken en exporteren van verkoopdata naar vendor-specifieke formaten voor AMD, Intel en Microsoft.

## 📋 Overzicht

Deze applicatie transformeert SAP/ICECAT verkooprapportage data naar de specifieke export formaten die vereist zijn door verschillende hardware vendors:
- **AMD** - AMD specifiek rapportage formaat
- **Intel** - Intel specifiek rapportage formaat  
- **Microsoft** - Microsoft OEM rapportage formaat

De tool filtert automatisch niet-relevante productcategorieën (zoals gaming peripherals en PC componenten) en formatteert de data volgens de exacte vereisten van elke vendor.

## 🚀 Functies

- **Drag & Drop Upload** - Sleep Excel bestanden direct naar de upload zone
- **Automatische Data Filtering** - Verwijdert automatisch uitgesloten productcategorieën
- **Multi-Vendor Export** - Genereert exports voor AMD, Intel en Microsoft vanuit één bronbestand
- **Intelligente Kolom Mapping** - Mapt automatisch bronkolommen naar vendor-specifieke velden
- **Product Beschrijving Generatie** - Combineert hardware specificaties tot gedetailleerde productbeschrijvingen
- **Datum Formaat Behoud** - Behoudt dd/mm/yyyy hh:mm:ss formaat voor invoice datums

## 💻 Installatie & Gebruik

### Vereisten
- Een moderne webbrowser (Chrome, Firefox, Safari, Edge)
- Excel bestanden in .xlsx of .xls formaat

### Stappen

1. **Open de applicatie**
   - Open `index.html` in je webbrowser
   - Geen installatie of server setup nodig!

2. **Upload je Excel bestand**
   - Sleep je SAP/ICECAT export bestand naar de upload zone
   - Of klik op "Choose File" om een bestand te selecteren

3. **Download de exports**
   - Na succesvolle verwerking verschijnen drie download knoppen
   - Klik op de gewenste export: AMD, Intel of Microsoft
   - Het bestand wordt automatisch gedownload met de juiste naamgeving

## 📊 Data Verwerking

### Uitgesloten Productcategorieën
De volgende productcategorieën worden automatisch gefilterd uit de exports:
- Muismatten
- Voedingen
- Game controllers
- Videokaarten
- Behuizingen
- Moederborden
- Desktop monitoren
- Processor koeling
- Handheld Consoles
- Headsets
- Netwerkadapters
- Toetsenbord en muis sets
- Muizen
- Toetsenborden

### Vendor-Specifieke Velden

#### AMD & Intel Export
- Partner ID wordt automatisch ingesteld op "COPACO"
- Distributor Branch wordt automatisch ingesteld op "NL"
- Product Description combineert: Processor info, Graphics cards, Operating System

#### Microsoft Export  
- Disti TPID wordt automatisch ingesteld op "201286"
- Bevat OEM specifieke velden zoals Operating System en Currency

## 📁 Project Structuur

```
copaco-dataverrijking/
├── index.html          # Hoofd HTML interface
├── app.js             # JavaScript logica voor data verwerking
├── images/            # Afbeeldingen en logo's
│   └── copaco_logo.webp
├── files/             # Voorbeeld export bestanden
│   ├── AMD.xlsx
│   ├── INTEL.xlsx
│   └── MICROSOFT.xlsx
└── README.md          # Deze documentatie
```

## 🔧 Technische Details

### Gebruikte Technologieën
- **HTML5** - User interface structuur
- **CSS3** - Styling en responsive design
- **JavaScript (ES6)** - Data verwerking logica
- **SheetJS (xlsx)** - Excel bestand lezen en schrijven

### Kolom Mapping
De applicatie gebruikt voorgedefinieerde mappings tussen brondata kolommen en vendor-specifieke export kolommen. Deze mappings zijn geconfigureerd in `exportMappings` object in app.js.

### Data Flow
1. Excel bestand wordt gelezen via FileReader API
2. Data wordt geconverteerd naar JSON formaat
3. Filtering wordt toegepast op basis van Web Hierarchy Description
4. Data wordt getransformeerd naar vendor-specifiek formaat
5. Nieuwe Excel bestanden worden gegenereerd voor download

## 🤝 Support

Voor vragen of problemen met de tool, neem contact op met het development team of maak een issue aan in de project repository.

## 📄 Licentie
© 2024–2025 Copaco Nederland B.V. – Alle rechten voorbehouden.  
Zie het bestand `LICENSE` voor details.
