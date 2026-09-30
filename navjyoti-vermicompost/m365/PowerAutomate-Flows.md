# Navjyoti alerts – Power Automate flows

Create each at **make.powerautomate.com → + Create**. Site address = your Navjyoti Production site.
Replace the Teams team/channel and email addresses with your own.

## 1. Overdue beds – daily 8:00

1. **Scheduled cloud flow** · every 1 day at 08:00 (time zone India Standard Time).
2. **SharePoint – Get items** · List: *Beds* · Filter Query:
   `Status ne 'Harvested' and ExpectedHarvest lt '@{formatDateTime(utcNow(),'yyyy-MM-dd')}'`
3. **Condition** · `length(body('Get_items')?['value'])` is greater than `0`.
4. If yes: **Data Operations – Create HTML table** · From: `body('Get_items')?['value']` · Custom columns: Batch = `item()?['Batch']?['Value']`, Bed = `item()?['BedNo']`, Expected = `item()?['ExpectedHarvest']`.
5. **Microsoft Teams – Post message in a chat or channel** · "⚠ Beds past expected harvest" + the HTML table.

## 2. Watering gap – daily 9:00

1. **Scheduled cloud flow** · daily 09:00.
2. **Get items** · List: *DailyLog* · Filter Query: `Activity eq 'Watering'` · Order By: `LogDate desc` · Top Count: `1`.
3. **Compose** (DaysSince):
   `if(empty(body('Get_items')?['value']), 99, div(sub(ticks(utcNow()), ticks(first(body('Get_items')?['value'])?['LogDate'])), 864000000000))`
4. **Condition** · `outputs('Compose')` is greater than `3` → **Post message** to the supervisor: "No watering logged for @{outputs('Compose')} days."

## 3. FG / EX lab report outside FCO – instant

1. **Automated cloud flow** · trigger **SharePoint – When an item is created** · List: *QualityControl*.
2. **Condition** (Advanced mode):
   ```
   @and(
     or(equals(triggerBody()?['Product']?['Value'],'FG'), equals(triggerBody()?['Product']?['Value'],'EX')),
     or(
       greater(float(coalesce(triggerBody()?['Moisture'],20)),25), less(float(coalesce(triggerBody()?['Moisture'],20)),15),
       less(float(coalesce(triggerBody()?['Nitrogen'],1)),1),
       less(float(coalesce(triggerBody()?['Phosphorus'],0.8)),0.8),
       less(float(coalesce(triggerBody()?['Potassium'],0.8)),0.8),
       greater(float(coalesce(triggerBody()?['CNRatio'],20)),20),
       less(float(coalesce(triggerBody()?['pH'],7)),6.5), greater(float(coalesce(triggerBody()?['pH'],7)),7.5)))
   ```
3. If yes: **Office 365 Outlook – Send an email (V2)** to the production manager:
   Subject `Lab report outside FCO – @{triggerBody()?['Product']?['Value']} @{triggerBody()?['ReportNo']}`,
   body listing Moisture, N, P, K, C:N, pH and a link `@{triggerBody()?['{Link}']}`.

## 4. Unpaid raw material – weekly Monday 10:00

1. **Scheduled cloud flow** · weekly, Monday 10:00.
2. **Get items** · *RawMaterial* · Filter Query: `PaymentStatus ne 'Paid'`.
3. **Condition** length > 0 → **Create HTML table** (Batch, RMLot, Supplier, QtyKg, Amount) → **Send an email** to accounts.

## Tips

- Test each flow with **Test → Manually** after creating a sample item.
- Filter queries use the column's internal name; the script created the lists with internal names equal to the display names shown here.
