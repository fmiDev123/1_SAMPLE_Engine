"""Draft market-dynamics input workbook for the Allergy Immunotherapy sample.

Read by the engine through REPORT_DYNAMICS (sheet names as the engine
expects).  Content is a first draft written for the sample and must be
validated by the research team before it is used in a paid report.
"""
import openpyxl

wb = openpyxl.Workbook()
wb.remove(wb.active)


def sheet(name, head, rows):
    ws = wb.create_sheet(name)
    ws.append(head)
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 40


TAKE = ("North America and Europe anchor the market through registered "
        "products and established prescribing, while Asia and the emerging "
        "regions add the fastest volume growth from a smaller base.")
sheet("Dynamics by Geography",
      ["Geography", "Region", "Growth Drivers", "Restraints",
       "Priority Opportunities", "Takeaway"], [
    ["", "North America",
     "Large diagnosed allergy population, strong allergist network, "
     "FDA-registered SLIT tablets and broad awareness of disease-modifying "
     "treatment",
     "Subcutaneous therapy needs supervised clinic visits; co-pay burden and "
     "variable payer coverage",
     "Home-based SLIT tablets, paediatric initiation and peanut oral "
     "immunotherapy", TAKE],
    ["", "Latin America",
     "Rising urban allergy prevalence and growing private-sector specialist "
     "care",
     "Limited reimbursement, uneven allergist coverage and import dependence "
     "for registered products",
     "Hospital and private-clinic programmes in Brazil and Mexico; locally "
     "produced extracts", ""],
    ["", "Europe",
     "Long-standing immunotherapy tradition, strong guideline support and "
     "tablet registrations across several allergen classes",
     "Pricing and tender pressure; heterogeneous national reimbursement",
     "Allergoid and molecular products, paediatric label extensions and "
     "tree-pollen indications", ""],
    ["", "East Asia",
     "Growing awareness of allergic rhinitis and asthma, high urban "
     "pollution exposure and expanding specialist care in China, Japan and "
     "South Korea",
     "Strict local registration, limited reimbursement outside Japan and a "
     "low share of patients on disease-modifying therapy",
     "Local registrations of house dust mite and cedar-pollen products; "
     "paediatric programmes", ""],
    ["", "South Asia & Pacific",
     "Large addressable allergic population, increasing private healthcare "
     "spending and Australasian allergen specialisation",
     "Low diagnosis rates, out-of-pocket payment and limited access to "
     "trained allergists",
     "Affordable standardised extracts, tele-allergy services and regional "
     "manufacturing", ""],
    ["", "Middle East & Africa",
     "Rising prevalence in Gulf countries and investment in specialist "
     "hospitals",
     "Limited local manufacturing, import registration lead times and few "
     "trained prescribers",
     "Specialist centres in the GCC and South Africa; distributor-led market "
     "entry", ""],
])


def kind(name, rows):
    sheet(name, ["Geography", "Title", "Description", "Period 1", "Period 2",
                 "Period 3", "Period 4"], [[""] + r for r in rows])


kind("Drivers", [
    ["Rising prevalence of allergic rhinitis, asthma and food allergy",
     "Allergic disease burden continues to climb with urbanisation, "
     "pollution exposure and longer pollen seasons. Immunotherapy is the "
     "only approach that targets the underlying immune response, so a "
     "larger diagnosed population feeds treatment-course demand.",
     "High", "High", "Medium", "Medium"],
    ["Shift from symptom relief to disease-modifying treatment",
     "Allergists increasingly position immunotherapy as a course that can "
     "change the disease, rather than lifelong antihistamine and "
     "corticosteroid use. Guideline support for rhinoconjunctivitis and "
     "allergic asthma reinforces the move.",
     "High", "High", "High", "Medium"],
    ["Registered SLIT tablets and home-based administration",
     "Sublingual tablets allow dosing at home after the first supervised "
     "dose. This widens access beyond specialist clinics and supports "
     "convenience, adherence and uptake among adults and children.",
     "Medium", "High", "High", "High"],
])
kind("Restraints", [
    ["Long treatment duration and adherence",
     "A course typically runs for several years, and discontinuation before "
     "completion limits the realised benefit and the revenue per patient.",
     "High", "High", "High", "Medium"],
    ["Safety management and supervised administration",
     "Injection-based therapy needs clinic supervision and preparedness for "
     "systemic reactions, which limits capacity and patient convenience.",
     "High", "High", "Medium", "Medium"],
    ["Reimbursement gaps and out-of-pocket cost",
     "Coverage varies by country and product, and patients may carry a large "
     "share of a multi-year course cost, especially in emerging markets.",
     "Medium", "Medium", "Medium", "Medium"],
])
kind("Opportunities", [
    ["Recombinant and molecular allergen preparations",
     "Better-defined, standardised preparations promise consistent potency "
     "and component-resolved treatment selection, opening a premium "
     "segment.", "Low", "Medium", "High", "High"],
    ["Emerging-market expansion in Asia Pacific and Latin America",
     "Growing private healthcare spending and increasing allergist capacity "
     "create room for registered, affordable immunotherapy outside the "
     "established markets.", "Medium", "Medium", "High", "High"],
    ["Paediatric and food-allergy immunotherapy",
     "Earlier intervention in children and registered oral immunotherapy for "
     "peanut allergy extend the treated population and the length of the "
     "commercial opportunity.", "Medium", "High", "High", "High"],
])
sheet("Driving Factors", ["Geography", "Side", "Factor", "Score", "Finding"], [
    ["", "Supply", "Allergen raw-material availability and quality", 6.5,
     "Pollen, mite and venom sources are seasonal and biologically "
     "variable, so secure and traceable sourcing is a differentiator."],
    ["", "Supply", "Manufacturing and potency standardisation", 7.5,
     "Few producers can run validated, standardised extract and tablet "
     "manufacturing at scale, which concentrates supply."],
    ["", "Supply", "Regulatory approval pathway", 7.0,
     "Product-level licensure raises the cost of entry but rewards "
     "registered, evidence-backed products."],
    ["", "Supply", "Cold-chain and distribution reach", 6.0,
     "Extracts and vaccines need controlled storage and specialist "
     "distribution into clinics and pharmacies."],
    ["", "Supply", "Pipeline and technology platforms", 7.5,
     "Allergoids, recombinant allergens and new tablets keep renewing the "
     "portfolio of leading players."],
    ["", "Demand", "Allergic disease prevalence", 9.0,
     "A rising allergic population is the primary volume driver across "
     "every region."],
    ["", "Demand", "Guideline and prescriber support", 7.5,
     "Specialist guidelines and prescriber familiarity with sublingual "
     "products raise initiation rates."],
    ["", "Demand", "Patient preference for home-based therapy", 7.0,
     "Convenience and fewer clinic visits support uptake and persistence."],
    ["", "Demand", "Paediatric and early-treatment uptake", 6.5,
     "Earlier intervention extends the treatment period per patient."],
    ["", "Demand", "Insurance and reimbursement coverage", 6.0,
     "Coverage determines whether patients complete a multi-year course."],
])
sheet("Five Forces", ["Geography", "Force", "Rating", "Evidence"], [
    ["", "Supplier Power", "Moderate",
     "Allergen source material and standardised-extract capability are "
     "specialised and held by a limited number of suppliers."],
    ["", "", "", "Biological variability and regulatory qualification of "
     "raw-material sources make supplier switching slow."],
    ["", "Buyer Power", "Moderate",
     "Hospitals, allergy clinics and specialty pharmacies purchase through "
     "tenders and distributors, which gives institutional buyers some "
     "leverage."],
    ["", "", "", "Prescribers, not patients, choose the therapy, and "
     "registered products limit easy substitution on price alone."],
    ["", "Competitive Rivalry", "Moderate",
     "A small set of specialist players leads the market, and competition "
     "centres on registered products, evidence and prescriber "
     "relationships."],
    ["", "", "", "Local and regional extract producers compete mainly on "
     "price and relationships within their home markets."],
    ["", "Threat of Substitutes", "Moderate",
     "Symptomatic pharmacotherapy and, in selected indications, biologic "
     "treatments offer alternatives to immunotherapy."],
    ["", "", "", "Immunotherapy keeps an advantage as a course that targets "
     "the underlying disease rather than controlling symptoms."],
    ["", "Threat of New Entrants", "Low",
     "Biologic manufacturing, potency standardisation and product-level "
     "licensure create high capital and regulatory barriers."],
    ["", "", "", "Clinical evidence requirements and prescriber trust take "
     "years to build."],
])
sheet("Recent Developments", ["Geography", "Year", "Development"], [
    ["", "2014", "The U.S. FDA approved the first sublingual immunotherapy "
     "tablets for grass and ragweed pollen allergy, opening the "
     "home-administration segment in North America."],
    ["", "2017", "The U.S. FDA approved the first sublingual tablet for "
     "house dust mite allergy."],
    ["", "2018", "Stallergenes and Greer Laboratories combined to form "
     "Stallergenes Greer, joining a European sublingual specialist with a "
     "U.S. extract manufacturer."],
    ["", "2019", "A tree-pollen sublingual tablet for birch-related allergy "
     "received European approval, widening tablet use in Europe."],
    ["", "2020", "The U.S. FDA approved the first oral immunotherapy for "
     "peanut allergy in children, creating a registered food-allergy "
     "immunotherapy category."],
    ["", "2020", "Nestlé Health Science completed its acquisition of the "
     "maker of the peanut oral immunotherapy product."],
])
wb.save(r"D:\PY\FMI_PPT\Output\AIT_Sample_work\AIT_Dynamics.xlsx")
print("saved")
