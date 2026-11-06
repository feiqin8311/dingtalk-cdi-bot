# Register Detail Merge Design

## Goal

Allow one shipment-detail SKU record to provide shared logistics metadata to multiple packing rows for the same shipment and SKU.

## Scope

Change only invoice-mode registration data assembly in `invoice_data.CustomsDeclaration.register_combined_invoice()`.

## Data Flow

The packing sheet remains the source of outbound-detail rows. The shipment-detail sheet supplies logistics metadata. A packing key may occur multiple times while its detail key occurs once; in that case each packing row receives the single detail record. When the detail sheet has duplicate keys, use a single deterministic detail record so the merge cannot multiply packing rows.

## Validation

The function must still raise its existing validation error when no shipment-detail row exists for a packing `SKU + 发货单号` key. The regression fixture models `T6HW84` as two packing rows with quantities 60 and 18 and one detail row with quantity 78.
