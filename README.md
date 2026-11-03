' ==============================================================================
' 1. DASHBOARD LOOKUPS & KPI CARDS (Sheet4 / Dashboard)
' ==============================================================================

' Fetch Product Unit Price dynamically based on Product ID in B1:
=XLOOKUP(B1, Dim_Biotherapeutic[product_id], Dim_Biotherapeutic[unit_price], "Not Found")

' Fetch Assigned Hub Name based on Shipment ID in B5:
=XLOOKUP(B5, Fact_Shipments[shipment_id], XLOOKUP(Fact_Shipments[hub_id], Dim_Distribution_Hub[Hub_ID], Dim_Distribution_Hub[Hub_Name]))

' Fetch Temperature Compliance Alert based on Shipment ID in B5:
=XLOOKUP(B5, Fact_Shipments[shipment_id], Fact_Shipments[temp_compliance])

' Fetch Delivery Status based on Shipment ID in B5:
=XLOOKUP(B5, Fact_Shipments[shipment_id], Fact_Shipments[delivery_status])

' Executive KPI Summary Cards:
Total Shipments (Cell D2)     : =COUNTA(Fact_Shipments[shipment_id])
OTIF Rate (Cell E2)           : =COUNTIF(Fact_Shipments[delivery_status], "Delivered-OTIF") / COUNTA(Fact_Shipments[shipment_id])
Compliance Rate (Cell F2)     : =COUNTIF(Fact_Shipments[temp_compliance], "Compliant") / COUNTA(Fact_Shipments[shipment_id])


' ==============================================================================
' 2. CONDITIONAL FORMATTING RULES (Fact_Shipments Sheet)
' ==============================================================================

' Rule 1: Temperature Excursion Alert (Red Fill)
Formula: =$F2="Excursion"

' Rule 2: Compliant Batch (Green Fill)
Formula: =$F2="Compliant"

' Rule 3: SLA Delivery Delay Alert (Light Red Fill)
Formula: =$E2="Delayed"
