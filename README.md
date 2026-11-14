# Transport Tracker

A trip and payment tracker for a small Indian transport business, built on **Google Sheets + Google Apps Script**. Staff enter trips and payments from a phone browser. The owner gets live reports in the sheet, PDF statements and one-tap WhatsApp payment reminders in Hindi and English.

No server, no app to install, no monthly cost: everything runs inside the owner's Google account.

**Jump to:** [Features](#features) · [Ideal daily flow (English)](#ideal-daily-flow) · [दैनिक उपयोग (हिंदी)](#दैनिक-उपयोग-का-तरीका-हिंदी) · [Setup](#setup-about-10-minutes) · [How it is built](#how-it-is-built)

<p>
<img src="screenshots/02-new-trip.png" width="240" alt="New trip form">
<img src="screenshots/07-summary.png" width="240" alt="Summary with WhatsApp reminder">
<img src="screenshots/08-whatsapp-message.png" width="240" alt="WhatsApp reminder message">
</p>

## Why it exists

Trips were tracked in a notebook and payments from memory. The goal was something drivers and office staff can use on a phone in under a minute per trip, that keeps names clean, and that shows at a glance which party owes how much.

## Features

**Mobile forms (open in any phone browser)**
- **New trip:** date (today by default), party, vehicle, up to 3 drivers, transporter, commodity, from/to, weight in quintal × rate → freight calculated automatically (or a fixed amount), advance. Only the basics are required; weight, rate and advance can be added later.
- **Payment:** against a specific trip or "on account", with an Expense field. Remaining = freight − advance − payments − expense.
- **Trip list:** view and edit trips, add or edit payments; filter by status, party, driver, vehicle, transporter, date range or text search. Cleared trips are hidden by default.
- **Summary:** totals by party, driver, vehicle or transporter for any period (last 7/30 days, this/last month, custom dates).
- **Masters:** Party, Vehicle, Driver, Transporter, with "+ Add new" from inside the Trip and Payment forms.
- Type-to-search pickers, so long lists stay easy to use.

**Clean data**
- Duplicate check ignores capital letters (Deepak = DEEPAK = deepak) and warns on similar names ("Did you mean…?") and repeat trips or payments.
- Names saved in proper case ("Sharma Traders"); vehicle numbers in ALL CAPS.
- Missing master entries created automatically from trips.
- Every record stores who entered it and who last changed it.
- Trip IDs like `T20261112-01`, payment IDs like `P20261112-01`.

**Reports and sharing**
- Sheet tabs: Dashboard, Trip Summary, Party / Vehicle / Driver / Monthly summaries, and a Custom Report with filters (dates, party, vehicle, driver, transporter).
- Print / Save as PDF from Trip list and Summary; party or transporter statement as PDF.
- **WhatsApp reminder** sent from the owner's own number with one tap (uses `wa.me` links, no paid API). Hindi first, then the same in English: date, route, vehicle, commodity, weight, rate, freight and pending amount for each trip.

**Safety**
- Nightly Excel backup to a Drive folder (keeps the last 60), optional weekly email copy.
- Daily pending-payment summary email and document-expiry alerts (insurance, permit, licence).
- Version number shown in the sheet menu and at the top of the phone forms.

---

## Ideal daily flow

### 1. One time: add your regular parties, vehicles and drivers
Open the form link and use the **Party**, **Vehicle**, **Driver** and **Transporter** tabs. You don't have to add everything up front: you can add a new name from inside the trip form when it first comes up.

<img src="screenshots/01-party-master.png" width="260" alt="Party master">

### 2. When a truck is loaded: enter the trip
1. Tap **New trip**. The date is already today.
2. Pick party, vehicle and driver. Type two or three letters to find a name (for example "ra" shows Ramesh and Rajesh).
3. If the name is new, tap **+ Add new**. The app warns you if a similar name already exists ("Did you mean Deepak?").
4. Pick from/to, enter weight (quintal) and rate. Freight is calculated for you. Use **Fixed amount** if the deal is a lump sum.
5. Enter any advance and tap **Save trip**. Weight, rate and advance can be left blank and filled in later.

<p>
<img src="screenshots/02-new-trip.png" width="240" alt="New trip">
<img src="screenshots/03-type-to-search.png" width="240" alt="Type to search">
<img src="screenshots/04-add-new-from-trip.png" width="240" alt="Add new driver from trip">
</p>

### 3. When money comes in: record the payment
1. Open **Trip list** and tap the trip (or search by party, vehicle or trip ID).
2. Tap **Add payment**. Enter the amount, mode (Cash / UPI / Bank / Cheque) and any expense (unloading, shortage, commission).
3. Tap **Save payment**. The trip's remaining amount updates straight away; once it reaches zero the trip moves to "Cleared".

<p>
<img src="screenshots/05-trip-list.png" width="240" alt="Trip list">
<img src="screenshots/06-payment.png" width="240" alt="Payment">
</p>

### 4. Every week: check who owes what and send reminders
1. Open **Summary**, choose **Party** and pick the party and period (for example **This month**).
2. You see trips, freight, advance, paid, expense and **remaining**.
3. Tap **WhatsApp reminder**. WhatsApp opens on your phone with the message ready, in Hindi and English. Check it and tap send.
4. For a formal statement, tap **View statement → Save as PDF** and send the PDF.

<p>
<img src="screenshots/07-summary.png" width="240" alt="Summary">
<img src="screenshots/08-whatsapp-message.png" width="240" alt="WhatsApp message">
<img src="screenshots/09-statement.png" width="240" alt="Statement">
</p>

### 5. Any time: find and fix
- **Trip list filters:** by driver, vehicle, transporter or date range, for example "all trips by Ramesh in the last 7 days".
- Tap **Edit trip** to correct a mistake. The record keeps who changed it.
- **Print / PDF** on Trip list and Summary for a paper copy.
- In the Google Sheet, the **🚚 Transport** menu refreshes reports, makes a backup now, and shows the code version.

<img src="screenshots/10-trip-list-filters.png" width="260" alt="Trip list filters">

**Who does what**

- **Driver or loading staff:** New trip only.
- **Office staff:** New trip, Payment, Trip list, corrections.
- **Owner:** Summary, WhatsApp reminders, statements, and the reports in the Google Sheet.

---

## दैनिक उपयोग का तरीका (हिंदी)

### 1. पहली बार: अपनी पार्टी, गाड़ी और ड्राइवर जोड़ें
फ़ॉर्म लिंक खोलें और **Party**, **Vehicle**, **Driver** और **Transporter** टैब में नाम जोड़ें। सब कुछ पहले से डालना ज़रूरी नहीं है, नया नाम ट्रिप भरते समय भी जोड़ सकते हैं।

<img src="screenshots/01-party-master.png" width="260" alt="पार्टी मास्टर">

### 2. गाड़ी लोड होने पर: ट्रिप भरें
1. **New trip** दबाएँ। तारीख़ अपने आप आज की होती है।
2. पार्टी, गाड़ी और ड्राइवर चुनें। नाम के दो-तीन अक्षर टाइप करें (जैसे "ra" लिखने पर Ramesh और Rajesh दिखेंगे)।
3. नया नाम हो तो **+ Add new** दबाएँ। मिलता-जुलता नाम पहले से हो तो ऐप पूछेगा ("Did you mean Deepak?")।
4. कहाँ से / कहाँ तक चुनें, वज़न (क्विंटल) और दर (rate) डालें। भाड़ा अपने आप बन जाएगा। तय रक़म का सौदा हो तो **Fixed amount** चुनें।
5. एडवांस डालें और **Save trip** दबाएँ। वज़न, दर और एडवांस बाद में भी भर सकते हैं।

<p>
<img src="screenshots/02-new-trip.png" width="240" alt="नई ट्रिप">
<img src="screenshots/03-type-to-search.png" width="240" alt="टाइप करके खोजें">
<img src="screenshots/04-add-new-from-trip.png" width="240" alt="ट्रिप से नया ड्राइवर जोड़ें">
</p>

### 3. पैसा आने पर: पेमेंट दर्ज करें
1. **Trip list** खोलें और ट्रिप पर टैप करें (या पार्टी, गाड़ी नंबर या ट्रिप नंबर से खोजें)।
2. **Add payment** दबाएँ। रक़म, तरीका (Cash / UPI / Bank / Cheque) और कोई खर्च (उतराई, कमी, कमीशन) डालें।
3. **Save payment** दबाएँ। ट्रिप का बाकी पैसा तुरंत अपडेट हो जाता है; बाकी शून्य होने पर ट्रिप "Cleared" हो जाती है।

<p>
<img src="screenshots/05-trip-list.png" width="240" alt="ट्रिप लिस्ट">
<img src="screenshots/06-payment.png" width="240" alt="पेमेंट">
</p>

### 4. हर हफ़्ते: किसका कितना बाकी है देखें और याद दिलाएँ
1. **Summary** खोलें, **Party** चुनें, फिर पार्टी और समय (जैसे **This month**) चुनें।
2. ट्रिप, भाड़ा, एडवांस, मिला हुआ पैसा, खर्च और **बाकी (Remaining)** दिखेगा।
3. **WhatsApp reminder** दबाएँ। आपके फ़ोन पर WhatsApp खुलेगा और मैसेज (हिंदी और अंग्रेज़ी में) तैयार होगा। देखकर भेज दें।
4. पक्का हिसाब भेजना हो तो **View statement → Save as PDF** दबाएँ और PDF भेजें।

<p>
<img src="screenshots/07-summary.png" width="240" alt="समरी">
<img src="screenshots/08-whatsapp-message.png" width="240" alt="WhatsApp मैसेज">
<img src="screenshots/09-statement.png" width="240" alt="स्टेटमेंट">
</p>

### 5. कभी भी: खोजें और सुधारें
- **Trip list के फ़िल्टर:** ड्राइवर, गाड़ी, ट्रांसपोर्टर या तारीख़ से, जैसे "पिछले 7 दिन में Ramesh की सारी ट्रिप"।
- ग़लती सुधारने के लिए **Edit trip** दबाएँ। किसने बदला, यह भी सेव रहता है।
- कागज़ पर कॉपी के लिए Trip list और Summary में **Print / PDF**।
- Google Sheet में **🚚 Transport** मेन्यू से रिपोर्ट रिफ़्रेश करें, तुरंत बैकअप लें और कोड का वर्ज़न देखें।

<img src="screenshots/10-trip-list-filters.png" width="260" alt="ट्रिप लिस्ट फ़िल्टर">

**कौन क्या करे**

- **ड्राइवर या लोडिंग स्टाफ़:** सिर्फ़ New trip।
- **ऑफ़िस स्टाफ़:** New trip, Payment, Trip list, सुधार।
- **मालिक:** Summary, WhatsApp रिमाइंडर, स्टेटमेंट और Google Sheet की रिपोर्ट।

---

## Setup (about 10 minutes)

1. Create a **blank Google Sheet**.
2. Open **Extensions → Apps Script**, delete what's there and paste the whole of [`Code.gs`](Code.gs).
3. Edit the `CONFIG` block near the top: business name, and optionally `OWNER_EMAIL` and `BACKUP_EMAIL` (leave blank to use your own account and skip the email backup).
4. Click **Save**, choose `setup` in the function list and click **Run**. Approve the permissions. If the log says "Setup paused", click Run again until it says "Setup complete".
5. **Deploy → New deployment → Web app**, Execute as **Me**, Who has access **Anyone**. Copy the web app link.
6. In Apps Script, choose `upgradeToMobileForms` in the function list and click **Run** once. This switches data entry to the mobile forms.
7. Back in the sheet, reload it. Use **🚚 Transport → Set mobile form link…** and paste the link. Open that link on your phone and add it to the home screen.

**Updating to a new version:** paste the new code, Save, then **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**, and reload the sheet. Check the version in the 🚚 Transport menu (last item) and at the top of the phone forms.

> **Privacy note:** anyone with the web app link can enter data, so share it only with your own staff.

## How it is built

- One file, `Code.gs` (Apps Script). The phone forms are an `HtmlService` web app served by `doGet()`, talking to the sheet through `google.script.run`.
- The Google Sheet is the database: Trips, Payments and master tabs hold the data; report tabs are rebuilt after each save.
- `LockService` keeps simultaneous saves safe; report rebuilding uses a separate lock so saving never waits on reports.
- Time-based triggers run the daily summary (8 am) and the nightly backup (10 pm), in the `Asia/Kolkata` timezone by default.
- Setup runs in resumable steps to stay under Apps Script's 6-minute limit.

Screenshots use sample data only.

## Version history

The full history is at the top of `Code.gs`. Current version: **v3.13**.
