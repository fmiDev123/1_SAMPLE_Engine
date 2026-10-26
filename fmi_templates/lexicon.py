"""
Industry lexicon for FMI auto-generated decks.

Every template that needs industry-specific wording uses an {lx_*} placeholder.
Pick the profile with ctx["industry"] = one of: food, consumer, healthcare, pharma, industrial.
Keys ending in _cap are generated automatically (first letter capitalised), so templates can
start a sentence with {lx_input_costs_cap}.

Rule: no em dashes or en dashes anywhere in this file (slide style rule).
"""

LEXICON = {
    # ------------------------------------------------------------------ FOOD & BEVERAGE
    "food": {
        "lx_end_user": "consumers and foodservice operators",
        "lx_channel_examples": "modern retail, convenience, foodservice and e-commerce channels",
        "lx_validation": "food-safety certification and clean-label credentials",
        "lx_innovation": "flavour, format and better-for-you reformulation",
        "lx_demand_effect": "widening consumption occasions and supporting repeat purchase",
        "lx_premium_driver": "premium ingredients and better-for-you claims",
        "lx_input_costs": "agricultural commodities, packaging and energy",
        "lx_regulator": "food-safety and labelling regulation",
        "lx_competitive_pressure": "private-label competition and retailer bargaining power",
        "lx_switching": "listing fees, shelf-space agreements and brand loyalty",
        "lx_substitutes": "adjacent snack, bakery and in-store fresh formats",
        "lx_entry_barriers": "brand investment, retail listings and distribution scale",
        "lx_localise": "formats, flavours and pack sizes",
        "lx_parent_market": "packaged food",
        "lx_factor10": "Trade Promotion Spend",
        "lx_factor10_why": "Retailer promotions and trade spend reduce net realised price, especially in mature grocery channels.",
        "lx_buyer_types": ["Grocery retail chain", "Foodservice operator", "Wholesale distributor", "Institutional caterer"],
        "lx_buyer_factors": ["Price and promotional support", "Food safety and quality", "Supply reliability", "Brand pull with shoppers"],
        "lx_vc": [
            ("Ingredient Suppliers", "Growers, millers and ingredient processors supply agricultural commodities and processed ingredients."),
            ("Manufacturers", "Industrial producers convert ingredients into finished, packaged products, which is the largest value-addition stage."),
            ("Distributors and Wholesalers", "Distributors manage ambient and chilled logistics, hold inventory and extend reach and credit to retail and foodservice accounts."),
            ("Retail and Foodservice", "Grocery, convenience, foodservice and online channels sell to consumers and capture point-of-sale margin."),
            ("Consumers", "Households and out-of-home consumers drive demand through taste, convenience and health preferences."),
        ],
        "lx_vc_top_reason": "processing, brand and scale advantages",
    },
    # ------------------------------------------------------------------ CONSUMER GOODS
    "consumer": {
        "lx_end_user": "consumers across income groups",
        "lx_channel_examples": "specialty retail, mass merchandise, direct-to-consumer and e-commerce channels",
        "lx_validation": "quality certification and brand trust",
        "lx_innovation": "product design, premium formats and sustainable materials",
        "lx_demand_effect": "widening usage occasions and supporting replacement demand",
        "lx_premium_driver": "premium materials and brand positioning",
        "lx_input_costs": "raw materials, packaging and logistics",
        "lx_regulator": "product-safety, labelling and sustainability regulation",
        "lx_competitive_pressure": "private-label competition and e-commerce price transparency",
        "lx_switching": "brand loyalty, retailer relationships and assortment agreements",
        "lx_substitutes": "lower-priced alternatives and adjacent product categories",
        "lx_entry_barriers": "brand building, retail distribution and marketing scale",
        "lx_localise": "formats, price points and pack sizes",
        "lx_parent_market": "consumer goods",
        "lx_factor10": "Marketing Spend",
        "lx_factor10_why": "Advertising, influencer and retailer trade spend shape net realised price and brand premium.",
        "lx_buyer_types": ["Mass retailer", "Specialty retailer", "E-commerce platform", "Distributor"],
        "lx_buyer_factors": ["Price and margin", "Brand demand", "Assortment breadth", "Fulfilment reliability"],
        "lx_vc": [
            ("Raw Material and Component Suppliers", "Suppliers provide raw materials, components and packaging inputs."),
            ("Brand Owners and Manufacturers", "Brand owners and contract manufacturers design, produce and market finished products, which is the largest value-addition stage."),
            ("Distributors and Wholesalers", "Distributors hold inventory and provide logistics, reach and credit to retail partners."),
            ("Retailers and E-commerce", "Store-based and online retailers sell to consumers and capture point-of-sale margin."),
            ("Consumers", "Consumers drive demand through brand preference, price sensitivity and lifestyle trends."),
        ],
        "lx_vc_top_reason": "brand, design and scale advantages",
    },
    # ------------------------------------------------------------------ HEALTHCARE / MEDICAL DEVICES / DIAGNOSTICS
    "healthcare": {
        "lx_end_user": "hospitals, clinics and diagnostic centres",
        "lx_channel_examples": "direct hospital sales, distributors, group purchasing organisations and online procurement",
        "lx_validation": "regulatory clearance and clinical evidence",
        "lx_innovation": "next-generation product design and digital integration",
        "lx_demand_effect": "expanding procedure volumes and clinical indications",
        "lx_premium_driver": "advanced features and better clinical outcomes",
        "lx_input_costs": "components, materials, sterilisation and compliance",
        "lx_regulator": "medical device regulation, including FDA and EU MDR requirements",
        "lx_competitive_pressure": "tender-based procurement and group purchasing power",
        "lx_switching": "clinician training, installed equipment and multi-year supply contracts",
        "lx_substitutes": "alternative procedures and competing technologies",
        "lx_entry_barriers": "regulatory clearance, clinical evidence and sales-force reach",
        "lx_localise": "product configurations and service models",
        "lx_parent_market": "medical technology",
        "lx_factor10": "Service Costs",
        "lx_factor10_why": "Installation, training and maintenance obligations add to lifetime cost and shape contract pricing.",
        "lx_buyer_types": ["Hospital", "Ambulatory or specialty clinic", "Diagnostic laboratory", "Group purchasing organisation"],
        "lx_buyer_factors": ["Clinical performance", "Total cost of ownership", "Regulatory compliance", "Service and training support"],
        "lx_vc": [
            ("Component and Material Suppliers", "Suppliers provide components, materials and sub-assemblies to validated specifications."),
            ("Manufacturers (OEMs)", "OEMs design, validate and produce finished products, which is the largest value-addition stage."),
            ("Distributors and GPOs", "Distributors and purchasing groups manage logistics, contracting and inventory for provider networks."),
            ("Healthcare Providers", "Hospitals, clinics and laboratories use the products to deliver care and drive procedure volumes."),
            ("Patients and Payers", "Patients and payers fund demand through reimbursement, insurance and out-of-pocket spending."),
        ],
        "lx_vc_top_reason": "regulatory know-how, intellectual property and scale advantages",
    },
    # ------------------------------------------------------------------ PHARMACEUTICALS
    "pharma": {
        "lx_end_user": "hospitals, pharmacies and treated patients",
        "lx_channel_examples": "hospital pharmacies, retail pharmacies and online pharmacies",
        "lx_validation": "regulatory approval, clinical data and reimbursement listing",
        "lx_innovation": "pipeline assets, new formulations and lifecycle management",
        "lx_demand_effect": "expanding treated patient populations",
        "lx_premium_driver": "differentiated efficacy and convenient dosing",
        "lx_input_costs": "active ingredients, excipients and manufacturing compliance",
        "lx_regulator": "drug regulation and pricing policy",
        "lx_competitive_pressure": "generic entry and payer price negotiation",
        "lx_switching": "prescriber familiarity, formulary position and patient continuity",
        "lx_substitutes": "generics, biosimilars and alternative therapies",
        "lx_entry_barriers": "R&D cost, regulatory approval and intellectual property protection",
        "lx_localise": "dosage forms and pack presentations",
        "lx_parent_market": "pharmaceutical",
        "lx_factor10": "Reimbursement Policy",
        "lx_factor10_why": "Reimbursement listings, reference pricing and payer rebates cap achievable net price.",
        "lx_buyer_types": ["Hospital pharmacy", "Retail pharmacy chain", "Wholesaler", "Government tender body"],
        "lx_buyer_factors": ["Clinical efficacy", "Net price and rebates", "Supply continuity", "Formulary status"],
        "lx_vc": [
            ("API and Excipient Suppliers", "Suppliers produce active ingredients and excipients under GMP conditions."),
            ("Manufacturers and CDMOs", "Innovators, generic makers and CDMOs formulate and produce finished doses, which is the largest value-addition stage."),
            ("Wholesalers and Distributors", "Wholesalers manage cold-chain and ambient logistics, inventory and credit for dispensing channels."),
            ("Pharmacies and Hospitals", "Pharmacies and hospitals dispense medicines to patients and capture dispensing margin."),
            ("Patients and Payers", "Patients and payers fund demand through reimbursement, insurance and out-of-pocket spending."),
        ],
        "lx_vc_top_reason": "intellectual property, regulatory exclusivity and scale advantages",
    },
    # ------------------------------------------------------------------ INDUSTRIAL / CHEMICALS / MATERIALS
    "industrial": {
        "lx_end_user": "OEMs and industrial end users",
        "lx_channel_examples": "direct OEM contracts, distributors and system integrators",
        "lx_validation": "product qualification and industry certification",
        "lx_innovation": "material performance, process efficiency and automation",
        "lx_demand_effect": "opening new application areas",
        "lx_premium_driver": "higher-performance grades and service bundles",
        "lx_input_costs": "raw materials, energy and freight",
        "lx_regulator": "environmental, safety and trade regulation",
        "lx_competitive_pressure": "low-cost imports and contract tendering",
        "lx_switching": "requalification, integration and retraining",
        "lx_substitutes": "alternative materials and technologies",
        "lx_entry_barriers": "capital intensity, qualification cycles and scale economics",
        "lx_localise": "product grades and service levels",
        "lx_parent_market": "industrial manufacturing",
        "lx_factor10": "Aftermarket Costs",
        "lx_factor10_why": "Spare parts, maintenance and technical service commitments influence contract pricing.",
        "lx_buyer_types": ["OEM", "Contract manufacturer", "Distributor", "End-use plant operator"],
        "lx_buyer_factors": ["Price and reliability", "Quality and compliance", "Lead time and service", "Technical support"],
        "lx_vc": [
            ("Raw Material Suppliers", "Primary producers supply base materials and intermediates."),
            ("Manufacturers and Processors", "Processors convert inputs into finished and semi-finished products, which is the largest value-addition stage."),
            ("Distributors and Integrators", "Distributors and system integrators provide logistics, inventory, credit and technical support."),
            ("OEMs and Contractors", "OEMs and contractors integrate the products into equipment, assemblies or projects."),
            ("End-Use Industries", "End-use industries drive demand through capital spending and production activity."),
        ],
        "lx_vc_top_reason": "processing know-how and scale advantages",
    },
}


def lexicon_for(industry: str) -> dict:
    """Return the lexicon dict with _cap and list-join helpers added."""
    if industry not in LEXICON:
        raise KeyError(f"Unknown industry '{industry}'. Use one of: {', '.join(LEXICON)}")
    lx = dict(LEXICON[industry])
    out = {}
    for k, v in lx.items():
        out[k] = v
        if isinstance(v, str):
            out[k + "_cap"] = v[:1].upper() + v[1:]
    for i, (name, desc) in enumerate(lx["lx_vc"], 1):
        out[f"lx_vc{i}"] = name
        out[f"lx_vc{i}_lc"] = name.lower()
        out[f"lx_vc{i}_desc"] = desc
    for i, b in enumerate(lx["lx_buyer_types"], 1):
        out[f"lx_buyer{i}"] = b
    for i, b in enumerate(lx["lx_buyer_factors"], 1):
        out[f"lx_buyer_factor{i}"] = b
    return out
