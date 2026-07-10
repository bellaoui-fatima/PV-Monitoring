import os
import time
import base64
from datetime import datetime, UTC
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import pandas as pd

import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException

load_dotenv()

os.makedirs("output", exist_ok=True)

USER = os.getenv("ENERGYSOFT_USER")
PASSWORD = os.getenv("ENERGYSOFT_PASSWORD")

BREVO_API_KEY = os.getenv("BREVO_API_KEY")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

OUTPUT_FILE = "anomalies_detectees.csv"
ATTACHMENT_NAME = "Rapport_Anomalies_Journalier.xlsx"

BASE_URL = "https://energysoft.app/odata/v4"
AUTH = HTTPBasicAuth(USER, PASSWORD)
HEADERS = {"accept": "application/json"}


def send_email_with_excel(subject, body, df_anomalies):
    """Génère un fichier Excel en mémoire et l'envoie en pièce jointe via Brevo."""
    configuration = sib_api_v3_sdk.Configuration()
    configuration.api_key['api-key'] = BREVO_API_KEY

    api_instance = sib_api_v3_sdk.TransactionalEmailsApi(
        sib_api_v3_sdk.ApiClient(configuration)
    )

    temp_excel_path = "temp_anomalies.xlsx"
    df_anomalies.to_excel(temp_excel_path, index=False, engine='openpyxl')

    with open(temp_excel_path, "rb") as f:
        excel_data = f.read()
        b64_content = base64.b64encode(excel_data).decode('utf-8')
    
    if os.path.exists(temp_excel_path):
        os.remove(temp_excel_path)

    attachment = sib_api_v3_sdk.SendSmtpEmailAttachment(
        content=b64_content,
        name=ATTACHMENT_NAME
    )

    email = sib_api_v3_sdk.SendSmtpEmail(
        sender={
            "name": "PV Monitoring",
            "email": SENDER_EMAIL
        },
        to=[{"email": RECEIVER_EMAIL}],
        subject=subject,
        text_content=body,
        attachment=[attachment]  
    )

    try:
        api_instance.send_transac_email(email)
        print("[EMAIL] Email envoyé avec succès avec le fichier Excel en pièce jointe.")
    except ApiException as e:
        print("[EMAIL] Erreur lors de l'envoi de l'e-mail :", e)


def execute_request_with_retry(url, params=None, method="GET"):
    """Exécute une requête HTTP et gère l'erreur 429 (Too many requests) de l'API."""
    while True:
        response = requests.get(url, auth=AUTH, headers=HEADERS, params=params)
        
        if response.status_code == 429:
            print("[API LIMIT] Limite de 30 req/min atteinte. Pause de 15 secondes...")
            time.sleep(15)
            continue
        return response


def get_all_sites():
    """Récupère dynamiquement la liste de toutes les centrales depuis Energysoft."""
    url = f"{BASE_URL}/Sites"
    response = execute_request_with_retry(url)
    if response.status_code != 200:
        print(f"[ERREUR] Impossible de charger les centrales. Status: {response.status_code}")
        return []
    return response.json().get("value", [])


def get_inverters(site_id):
    """Récupère tous les onduleurs d'un site spécifique."""
    url = f"{BASE_URL}/Sites('{site_id}')/Inverters"
    response = execute_request_with_retry(url)
    if response.status_code != 200:
        return []
    return response.json().get("value", [])


def get_instant_power_measure(inverter_id, date_str):
    """Récupère la dernière mesure de puissance d'un onduleur pour la journée en cours."""
    url = f"{BASE_URL}/Inverters({inverter_id})/Measures"
    query_string = f"$top=5&$filter=MeasureType eq 'power' and date ge {date_str}"
    full_url = f"{url}?{query_string}"
    
    response = execute_request_with_retry(full_url)
    if response.status_code != 200:
        return None
    
    measures = response.json().get("value", [])
    if not measures:
        return None
        
    measures_sorted = sorted(
        measures, 
        key=lambda x: x.get("timestamp", ""), 
        reverse=True
    )
    return measures_sorted[0]


def main():
    if not all([USER, PASSWORD, BREVO_API_KEY, SENDER_EMAIL, RECEIVER_EMAIL]):
        print("Erreur : Un ou plusieurs identifiants sont manquants dans le fichier .env.")
        return

    # Capture de la date et de l'heure exacte de l'appel API
    today_str = datetime.now(UTC).strftime("%Y-%m-%d")
    execution_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"--- DÉMARRAGE DE LA SUPERVISION EN TEMPS RÉEL (100% API) ---")
    print(f"Date d'analyse : {today_str} à {execution_time_str}")
    print("[API] Récupération de la liste des centrales...")
    
    # Étape 1 : Récupération automatique de toutes les centrales
    all_sites = get_all_sites()
    if not all_sites:
        print("[FIN] Aucun site trouvé ou erreur d'authentification.")
        return

    print(f"Nombre de centrales détectées sur le compte : {len(all_sites)}\n")

    plant_anomaly_table = []
    total_inverters_analysed = 0  

    for site in all_sites:
        site_id = site.get("ID")
        site_name = site.get("Name", "Nom Inconnu")
        
        if not site_id:
            continue

        print(f"\n[CENTRALE] Analyse du site : {site_name} (ID: {site_id})...")
        
        inverters = get_inverters(site_id)
        time.sleep(0.5)

        if not inverters:
            print("  -> Aucun onduleur trouvé sur cette centrale.")
            continue

        for inverter in inverters:
            total_inverters_analysed += 1  
            inv_id = inverter.get("ID")
            inv_name = inverter.get("Name")
            
            latest_measure = get_instant_power_measure(inv_id, today_str)
            
            if latest_measure:
                power_value = latest_measure.get("Value")
                timestamp_raw = (
                    latest_measure.get("timestamp", "")
                    .replace("T", " ")
                    .replace("Z", "")
                )
                
                if power_value == 0.0 or power_value == 0:
                    print(f"    [ANOMALIE] '{inv_name}' à {power_value} kW à {timestamp_raw} -> CAPTURÉ")
                    
                    plant_anomaly_table.append({
                        "Date_Appel": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Site_ID": site_id,
                        "Site_Name": site_name,
                        "Inverter_ID": inv_id,
                        "Inverter_Name": inv_name,
                        "Observation_Time": timestamp_raw,
                        "Power_Value": power_value,
                        "Status": "Anomaly (Power = 0)"
                    })
                else:
                    print(f"    [SKIP] '{inv_name}' en production : {power_value} kW")
            else:
                print(f"    [ANOMALIE] Perte de communication pour '{inv_name}' -> CAPTURÉ")
                plant_anomaly_table.append({
                    "Date_Appel": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Site_ID": site_id,
                    "Site_Name": site_name,
                    "Inverter_ID": inv_id,
                    "Inverter_Name": inv_name,
                    "Observation_Time": "N/A",
                    "Power_Value": "N/A",
                    "Status": "Communication Loss"
                })
            
            time.sleep(0.5)

    # ==========================
    # Statistiques globales
    # ==========================
    total_plants = len(all_sites)
    power_zero_count = sum(1 for a in plant_anomaly_table if a["Status"] == "Anomaly (Power = 0)")
    communication_loss_count = sum(1 for a in plant_anomaly_table if a["Status"] == "Communication Loss")
    total_anomalies = len(plant_anomaly_table)
    plants_with_anomalies = len(set(a["Site_ID"] for a in plant_anomaly_table))
    plants_without_anomalies = total_plants - plants_with_anomalies
    affected_percentage = (plants_with_anomalies / total_plants * 100) if total_plants else 0

    print("\nSTATISTIQUES FINALES")
    print("-" * 50)
    print(f"Total centrales analysées   : {total_plants}")
    print(f"Total onduleurs analysés    : {total_inverters_analysed}")
    print(f"Anomalies Puissance = 0     : {power_zero_count}")
    print(f"Pertes de communication     : {communication_loss_count}")
    print(f"Total anomalies détectées   : {total_anomalies}")
    print(f"Centrales affectées         : {plants_with_anomalies}/{total_plants} ({affected_percentage:.2f}%)")

    print("\n" + "=" * 60)
    print("ANALYSE TERMINÉE :")

    if plant_anomaly_table:
        df_new_anomalies = pd.DataFrame(plant_anomaly_table)
        
        if os.path.exists(OUTPUT_FILE):
            df_new_anomalies.to_csv(OUTPUT_FILE, mode='a', header=False, index=False, encoding='utf-8')
            print(f"[STOCKAGE] {len(df_new_anomalies)} anomalie(s) ajoutée(s) dans '{OUTPUT_FILE}'")
        else:
            df_new_anomalies.to_csv(OUTPUT_FILE, mode='w', header=True, index=False, encoding='utf-8')
            print(f"[STOCKAGE] Nouveau fichier historique créé '{OUTPUT_FILE}'")
            
        print("\n[NOTIFICATION] Préparation de l'e-mail avec fichier joint...")
        
        # Ajout de l'heure d'exécution dans le corps du mail
        body_message = (
            f"Bonjour,\n\n"
            f"Une ou plusieurs anomalies ont été détectées sur votre parc de centrales photovoltaïques lors de l'analyse du {today_str}.\n"
            f"Heure exacte de l'appel API : {execution_time_str}\n\n"
            f"Vous trouverez en pièce jointe le fichier Excel contenant la liste complète des {len(df_new_anomalies)} onduleurs en anomalie.\n\n"
            f"RÉSUMÉ DES STATISTIQUES\n"
            f"-----------------------------\n"
            f"Total Centrales : {total_plants}\n"
            f"Total Onduleurs Analysés : {total_inverters_analysed}\n"
            f"Anomalies Power=0 : {power_zero_count}\n"
            f"Pertes de Communication : {communication_loss_count}\n"
            f"Total Anomalies : {total_anomalies}\n"
            f"Taux d'impact parcs : {plants_with_anomalies}/{total_plants} ({affected_percentage:.2f}%)\n\n"
            f"Détails des alertes :\n"
        )

        for anomaly in plant_anomaly_table:
            body_message += f"- Centrale : {anomaly['Site_Name']} | Onduleur : {anomaly['Inverter_Name']} ({anomaly['Power_Value']} kW à {anomaly['Observation_Time']}) - Statut : {anomaly['Status']}\n"

        body_message += "\nCordialement,\nSystème de Supervision Automatique"

        send_email_with_excel(
            subject=f"Alerte de Supervision : {len(df_new_anomalies)} anomalies détectées ({today_str})", 
            body=body_message, 
            df_anomalies=df_new_anomalies
        )
    else:
        print("Aucune anomalie détectée sur l'ensemble de vos centrales.")
    
    print("=" * 60)


if __name__ == "__main__":
    main()