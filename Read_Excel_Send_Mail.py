import openpyxl
import win32com.client as win32
import pandas as pd


sender_email = "Enter your Email hier"


excelFile = pd.read_excel('List.xlsx', sheet_name='Tabelle1')

for index, row in excelFile.iterrows():
    inventarnummer = row['Inventarnummer']
    name = row['Name']
    email = row['E-Mail']

    # Vorlage für die Nachricht
    ausgabe = f"Hallo {name},\n\n" \
              f"wir freuen uns Ihnen mitteilen zu können, dass es Zeit ist Ihr altes Notebook gegen ein leistungsfähigeres Modell auszutauschen.\n" \
              f"Buchen Sie dafür einen Termin für den Austausch Ihres Geräts {inventarnummer} bis spätestens Datum bequem über den folgenden Link: [Link].\n" \
              f"Beachten Sie, dass der Link nur für eine begrenzte Zeit gültig ist. Falls Sie Unterstützung bei der Terminbuchung benötigen oder zu dem Austausch Fragen haben, stehen wir oder Ihre Lokale IT Ihnen gerne zur Verfügung.\n" \
              f"\n" \
              f"Am Ende dieser E-Mail finden Sie Antworten zu den häufigsten gestellten Fragen in Form eines FAQ´s.\n" \
              f"Ebenfalls hat Ihnen die IT-Abteilung ein E-Mail gesendet, in dem Sie auch hilfreiche Informationen finden können.\n" \
              f"\n" \
              f"Wir freuen uns auf unseren Termin.\n" \
              f"\n" \
              f"Mit freundlichen Grüßen,\n" \
              f"Clientservice-Team von Bechtle\n\n" \
              f"FAQ\n" \
              f"F: Darf ich das alte Netzteil mit dem neuen Notebook verwenden?\n" \
              f"A: Da sich der Stromstecker des Netzteils geändert hat, erhalten Sie beim Notebooktausch ein neues USB-C Netzteil. Bereiten Sie das alte Netzteil vor, sodass dieses beim Austausch mitgenommen werden kann.\n" \
              f"F: Welches Gerät werde ich erhalten?\n" \
              f"A: Die Geräte wurden von der IT zusammengestellt und 1 zu 1 getauscht. Heißt, Reisenotebooks gegen Reisennotebooks und Standardnotebooks gegen Standardnotebooks Weitere Informationen zu den technischen Details und Modellen finden Sie unter folgendem Link: IT Devices\n" \
              f"F: Kann ich ein anderes Modell oder Gerätetyp erhalten?\n" \
              f"A: Unter bestimmten Umständen ist das möglich. Erstellen Sie hierfür ein internes Ticket über das IT Support Portal. Hier der Link zum: IT Support Portal\n" \
              f"F: Ich benötige zusätzliches Zubehör. Wo kann ich dieses finden?\n" \
              f"A: Das Zubehör wird Ihnen von der Werkzeugausgabe bereitgestellt. Weitere Informationen finden Sie unter folgendem Link: Werkzeugausgabe\n" \
              f"F: Ich kann mein Gerät nicht direkt nach dem Austausch abgeben, da ich mich gerade in einem wichtigen Projekt befinde.\n" \
              f"A: Sie können Ihr Gerät bis max. 2 Wochen nach dem Austausch behalten, bevor Sie dieses bei Ihrer lokalen IT abgeben müssen.\n" \
              f"F: Ich kann gerade keine freien Termine im Tool finden. Wie komme ich an meinen Termin?\n" \
              f"A: Wenn Sie keinen auswählbaren Termin finden, sind entweder schon alle Termine in dem angegebenen Zeitraum vergeben oder Sie müssen auf einen neue Terminvergabe warten.\n" \
              f"F: Mir ist etwas dazwischenkommen und kann somit meinen gebuchten Termin nicht wahrnehmen. Kann ich diesen Termin auch verschieben oder absagen?\n" \
              f"A: Falls Sie an einem Termin nicht teilnehmen können, bitten wir Sie den Termin über das Tool abzusagen, damit ein anderer Mitarbeiter diesen Termin buchen kann.\n" \
              f"F: Der angegebene Rechnername bzw. die angegebene Inventarnummer entspricht nicht meinem Rechner (Desktop / Notebook / Workstation).\n" \
              f"A: Bitte prüfen Sie, ob es sich eventuell um einen Abteilungsrechner, Besprechungsrechner oder sonstigen Rechner in Ihrer Abteilung handelt, an dem Sie vor kurzem angemeldet waren. Bitte leiten Sie in diesem Fall diese E-Mail zur Terminvereinbarung an die für diesen Rechner zuständige Person/Sekretariat weiter oder informieren Sie uns, dass dieses nicht Ihr Rechner ist.\n"

    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = "Austausch Ihres Notebooks"
        mail.Body = ausgabe
        
        mail.SentOnBehalfOfName = sender_email

        mail.Send()
        print(f"E-Mail an {email} wurde erfolgreich versendet.")
    except Exception as e:
        print(f"Fehler beim Senden der E-Mail an {email}: {str(e)}")

