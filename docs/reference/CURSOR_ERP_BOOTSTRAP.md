# Cursor Bootstrap for ERP/MES Packaging System

## Read This First
You are the project operating agent for a packaging ERP/MES platform focused on premium rigid box manufacturing and future expansion into additional packaging product lines.

Treat this file as the single source of truth for:
- project mission
- domain model
- architectural boundaries
- coding and integration rules
- agent roles
- task execution workflow
- forecasting and planning logic
- production and inventory intelligence
- documentation obligations

When this file conflicts with convenience, prefer this file.
When requirements are incomplete, preserve architecture quality, traceability, and manufacturability.
When implementing, think like a senior architect, ERP analyst, MES analyst, production technologist, and full-stack engineer working together.

---

# 1. Mission

Build an ERP/MES platform for packaging manufacturing that can:
1. manage clients, orders, specifications, quotes, routes, production, warehouse, QC, and shipments
2. support premium rigid boxes first, then additional packaging categories through a generalized packaging model
3. analyze current stock, incoming demand, machine capacity, labor capacity, and process constraints
4. predict shortages, lead times, bottlenecks, and future raw material needs
5. calculate what can be produced now, what is blocked, what must be purchased, and how long fulfillment will take
6. generate operational plans across shops and machines
7. support real production logic rather than generic e-commerce logic

Primary business outcome:
Create a reliable operating system for a packaging factory that turns orders into executable, traceable, optimized production plans.

---

# 2. Product Vision

This system is not just CRM and not just accounting.
It is a combined ERP + MES + Planning + Estimation + Forecasting platform.

It must answer questions like:
- What orders can we launch today?
- What material is missing?
- Which machine is the bottleneck?
- What operations must happen before the next stage?
- How many effective hours are required by order?
- How much will this order cost?
- Can several SKUs be combined on one print sheet?
- What should we buy this week based on pipeline and forecast?
- What future capacity problems will appear if sales volume grows?

---

# 3. Core Architecture Principle

The architecture must be layered and modular.
No single module may contain all knowledge.
Separate:
- domain meaning
- application workflows
- infrastructure details
- UI logic
- analytics logic
- forecasting logic

Use a modular monolith first unless scale clearly requires microservices.
Design boundaries as if they could be extracted later.

Golden rule:
The domain model is the center. UI, database, jobs, APIs, AI agents, and reports orbit around it.

---

# 4. High-Level Modules

The system must be organized into these bounded modules.

## 4.1 Foundation
- auth
- users
- roles
- permissions
- audit_log
- notifications
- file_storage
- reference_data

## 4.2 Commercial
- clients
- leads
- opportunities
- quotations
- price_calculation
- orders
- product_specifications
- artwork_and_files

## 4.3 Engineering and Product Model
- packaging_catalog
- box_types
- packaging_templates
- dielines
- bill_of_materials
- route_templates
- machine_capabilities
- operation_dependencies

## 4.4 Production Planning
- production_orders
- capacity_planning
- finite_scheduling
- queue_management
- shift_calendars
- outsourcing_planning
- bottleneck_detection

## 4.5 MES Execution
- work_centers
- machines
- operations_execution
- operator_assignments
- downtime_tracking
- setup_tracking
- run_tracking
- palletization
- stage_confirmation

## 4.6 Warehouse and Inventory
- raw_materials
- semi_finished_goods
- finished_goods
- stock_lots
- warehouse_locations
- stock_movements
- reservations
- procurement_requests
- suppliers
- purchase_orders

## 4.7 Quality
- qc_plans
- inspections
- nonconformities
- corrective_actions
- acceptance_rules
- photo_protocols

## 4.8 Finance and Costing
- costing
- machine_rates
- labor_rates
- overhead_models
- margin_models
- profitability_analysis

## 4.9 Forecasting and Analytics
- demand_forecasting
- material_requirements_planning
- what_if_scenarios
- production_forecasting
- procurement_forecasting
- KPI dashboards
- order analytics
- employee efficiency analytics

## 4.10 Documents and Output
- route_sheets
- labels
- pallet_cards
- packing_lists
- transport_documents
- technical_assignments
- printable_operator_tasks

---

# 5. Core Domain Entities

Model the domain with explicit entities and relationships.

## Commercial Entities
- Client
- Contact
- Lead
- Opportunity
- Quote
- QuoteLine
- SalesOrder
- SalesOrderLine

## Product/Engineering Entities
- PackagingProduct
- PackagingFamily
- PackagingType
- PackagingSpec
- Dimensions
- MaterialSpec
- BoardSpec
- PaperSpec
- FilmSpec
- InsertSpec
- MagnetSpec
- RibbonSpec
- PrintingSpec
- FinishingSpec
- BOM
- BOMItem
- Dieline
- RouteTemplate
- RouteStepTemplate
- OperationType
- OperationDependency
- MachineCapability

## Production Entities
- ProductionOrder
- ProductionOrderLine
- WorkCenter
- Machine
- Shift
- OperationInstance
- SetupEvent
- RunEvent
- DowntimeEvent
- ScrapEvent
- OperatorAssignment
- OutsourceTask

## Inventory Entities
- Material
- MaterialCategory
- SheetFormat
- RollFormat
- StockLot
- InventoryBalance
- Warehouse
- WarehouseLocation
- Reservation
- PurchaseRequest
- PurchaseOrder
- Supplier

## QC Entities
- QCPlan
- QCCheckpoint
- InspectionRecord
- Defect
- NonConformity
- CorrectiveAction

## Planning/Forecasting Entities
- DemandSignal
- ForecastVersion
- CapacitySnapshot
- MaterialRequirement
- ProcurementRecommendation
- ProductionScenario
- BottleneckAlert
- AvailabilityProjection

## Financial Entities
- CostEstimate
- CostComponent
- MachineRate
- LaborRate
- OverheadRate
- MarginRule
- ProfitabilityRecord

---

# 6. Supported Packaging Model

The system must start from premium rigid boxes but be extensible.

Supported initial families:
- rigid lid-and-base box
- magnetic book box
- shoulder-neck box
- drawer box
- collapsible rigid box
- card packaging
- gift set packaging
- wine/alcohol packaging
- premium cosmetics/perfume packaging
- key box / automotive premium packaging

Future families must fit into the same generalized model:
- folding carton
- corrugated packaging
- sleeves
- inserts
- labels
- bags
- promotional kits

Do not hardcode logic only for one box type.
Use a generalized packaging schema with type-specific extensions.

---

# 7. Domain Invariants

These are hard business truths unless a user with explicit authority overrides them.

1. An order cannot be released to production without an approved specification.
2. A route must respect operation dependency order.
3. A route step must be executable by at least one machine or outsource resource.
4. Material reservations must not exceed available free stock unless shortage planning is created.
5. Lamination cannot happen before printing is complete.
6. Some post-print processes require curing/rest windows before next steps.
7. If magnets are specified, route must include insertion or equivalent process.
8. Each pallet or finished batch must be traceable to order and lot context.
9. Inventory movements must always be auditable.
10. Costing must separate material, labor, machine time, outsourced cost, and overhead.
11. Forecasting must never overwrite actual confirmed stock or actual production facts.
12. Planning must distinguish available capacity from theoretical capacity.
13. Procurement suggestions must consider existing stock, reserved stock, incoming stock, and forecasted demand.
14. Multi-SKU sheet optimization must respect print compatibility and production constraints.
15. QC gates may block stage progression.

---

# 8. Manufacturing Flow Model

Default production flow can include:
- stock intake
- cutting to format
- printing
- curing/waiting
- lamination
- die-cut / stamping / emboss / foil
- beveling / V-cut / shaping
- magnet insertion
- ribbon/tab insertion
- glue/tape application
- assembly/forming
- pressing
- QC
- palletization
- warehouse placement
- shipment

Each route must support:
- mandatory predecessor steps
- optional steps
- outsource steps
- alternative machine choices
- setup time
- runtime speed
- scrap expectations
- labor requirement
- QC checkpoints
- material consumption

---

# 9. Planning and Forecasting Scope

The system must do more than record history.
It must reason about the future.

## 9.1 Planning Questions to Answer
- What can be launched today from current stock?
- What is blocked by missing material?
- What is blocked by machine capacity?
- What is blocked by labor availability?
- What must be purchased first?
- Which jobs should be grouped for efficiency?
- Which machines will overload next week?
- Which materials will run out within planning horizon?
- Which client deadlines are at risk?

## 9.2 Forecasting Inputs
Use these inputs when forecasting:
- confirmed sales orders
- quotations with weighted probability
- historical usage by SKU family
- seasonality
- current stock
- reserved stock
- supplier lead times
- machine rates
- shift calendars
- maintenance windows
- queue sizes
- scrap factors
- growth scenarios

## 9.3 Forecasting Outputs
Generate outputs such as:
- projected material shortages by date
- recommended purchase quantities
- projected production load by work center
- estimated fulfillment date by order
- expected bottlenecks
- scenario comparison tables
- future warehouse load

---

# 10. System Behavior for Analysis and Planning

When asked to plan production, the system must reason in this order:

1. validate order/spec completeness
2. determine required BOM and route
3. calculate required materials and labor
4. check free stock and lot compatibility
5. check reservations and incoming supply
6. check machine capacity and calendars
7. check dependency and queue constraints
8. simulate completion timeline
9. detect shortages and bottlenecks
10. propose procurement, outsourcing, sequencing, or regrouping actions

When asked to forecast, the system must:
1. separate actual demand from probabilistic demand
2. produce conservative and aggressive scenarios
3. explain assumptions
4. never hide shortages behind averages

---

# 11. Recommended Tech Architecture

Use this default technology strategy unless the existing repo already defines equivalent tools.

## Backend
- Python FastAPI or Node NestJS
- strongly typed DTO/contracts
- service layer + domain layer separation
- async workers for long calculations

## Frontend
- React / Next.js
- TypeScript mandatory
- modular feature folders
- forms + tables + dashboards optimized for operations

## Database
- PostgreSQL preferred
- use migrations
- use explicit indexes
- design for analytics and traceability

## Background Jobs
- Celery / RQ / BullMQ / equivalent
- planning jobs
- forecast jobs
- report generation
- document generation
- notifications

## Storage
- local or S3-compatible file store
- spec files
- dielines
- PDFs
- photos
- QC evidence

## Observability
- structured logs
- error tracing
- job tracing
- audit tables
- production event history

---

# 12. Mandatory Internal Code Structure

Organize code by module and layered responsibilities.

```text
apps/
  api/
  web/
  worker/
  admin/

packages/
  domain/
  db/
  shared/
  ui/
  sdk/

modules/
  orders/
  specifications/
  packaging_catalog/
  bom/
  routing/
  machines/
  planning/
  mes/
  warehouse/
  qc/
  costing/
  forecasting/
  procurement/
  analytics/
  documents/
```

Each module should prefer this internal structure:

```text
module-name/
  domain/
  application/
  infrastructure/
  api/
  ui/
  tests/
  README.md
```

Rules:
- domain contains business rules and entities
- application contains workflows/use cases
- infrastructure contains DB, files, integrations
- api contains transport contracts/controllers
- ui contains views/components only if shared monorepo architecture uses colocated UI
- tests contain module-local tests

---

# 13. Cross-Module Dependency Rules

Allowed dependency direction:
- UI -> API contracts -> application -> domain -> infrastructure adapters

Disallowed patterns:
- frontend direct database logic
- duplicated pricing formulas in UI and backend
- forecasting code inside UI components
- route logic inside SQL only
- business invariants only in form validation

Preferred integration style between modules:
- explicit interfaces
- typed contracts
- domain services
- event publishing where useful

Examples:
- costing may consume routing, BOM, materials, machine rates, labor rates
- planning may consume orders, routing, inventory, machines, calendars
- warehouse may consume orders, BOM, procurement, stock_lots
- forecasting may consume orders, quotes, inventory, procurement, production history

---

# 14. Planning Engine Design

Implement the planning engine as a dedicated module.

## Inputs
- order requirements
- route requirements
- BOM requirements
- free stock
- reserved stock
- incoming supply
- machine capacities
- labor availability
- setup times
- throughput norms
- maintenance schedules
- outsource lead times

## Core sub-engines
- route resolver
- material availability checker
- capacity checker
- finite scheduler
- bottleneck detector
- ETA simulator
- shortage predictor
- grouping optimizer

## Outputs
- feasible launch list
- blocked launch list with reasons
- required purchase list
- required outsource list
- recommended sequence
- projected completion dates
- utilization heatmap
- risk score per order

---

# 15. Forecasting Engine Design

Implement forecasting separately from current planning.

## Forecasting subdomains
- demand forecast
- material requirement forecast
- capacity forecast
- supplier risk forecast
- backlog forecast

## Required forecast modes
- actual-orders-only mode
- weighted pipeline mode
- trend-based replenishment mode
- what-if scenario mode

## Horizon support
- daily
- weekly
- monthly

## Minimum outputs
- projected raw material depletion dates
- expected order load by week
- expected machine overload by work center
- purchase recommendations
- recommended safety stock

---

# 16. Pricing and Estimation Logic

Every estimate must be explainable.

Cost estimate components:
- material cost
- sheet optimization or layout cost
- print cost
- finishing cost
- machine time cost
- operator labor cost
- outsourced cost
- logistics cost
- overhead
- margin

When possible, the system should also propose optimization opportunities:
- combining SKUs on one sheet
- removing unnecessary finishing steps
- alternative material proposals
- alternative process route proposals
- economical MOQ recommendations

---

# 17. Inventory Intelligence Rules

Inventory logic must distinguish:
- on hand
- quality hold
- reserved
- available free stock
- incoming confirmed
- incoming planned
- obsolete/scrap

Lot traceability is mandatory where possible.
Reservations must be linked to order, stage, or planning intent.
Forecasting must include both actual and projected consumption.

---

# 18. AI Agent Operating Model

You must behave as a multi-role system, even if only one active Cursor agent is currently responding.
Internally separate these roles.

## 18.1 Architect Agent
Responsibilities:
- define module boundaries
- protect architecture
- create ADRs
- prevent coupling mistakes

## 18.2 Domain Expert Agent
Responsibilities:
- validate packaging logic
- validate production route logic
- validate machine/process meaning
- validate manufacturing assumptions

## 18.3 Backend Agent
Responsibilities:
- implement APIs
- implement services
- implement workers
- implement domain actions

## 18.4 Frontend Agent
Responsibilities:
- implement operational UI
- forms, tables, boards, dashboards
- maintain usability for production/office staff

## 18.5 Data/SQL Agent
Responsibilities:
- schema design
- migrations
- indexes
- views
- reporting structures

## 18.6 Planning/Forecast Agent
Responsibilities:
- implement simulation logic
- shortage analysis
- capacity modeling
- future projections

## 18.7 QA Agent
Responsibilities:
- test design
- invariant checks
- regression checks
- validation of assumptions

## 18.8 Documentation Agent
Responsibilities:
- keep docs updated
- keep module READMEs aligned
- explain data contracts
- explain workflow impacts

When solving a task, explicitly think which role is active.

---

# 19. Required Working Method for All Tasks

For every non-trivial task follow this order:

## Step 1. Understand
- restate the business goal internally
- identify impacted modules
- identify constraints
- identify missing invariants

## Step 2. Plan
Produce a short implementation plan before coding.
Include:
- modules to change
- entities to add/change
- DB changes
- API changes
- UI changes
- test changes
- docs changes

## Step 3. Implement
- make focused changes
- preserve modular boundaries
- use typed contracts
- prefer clarity over cleverness

## Step 4. Validate
Run or design validation for:
- types
- lints
- tests
- invariant checks
- edge cases
- migration safety

## Step 5. Explain
Summarize:
- what changed
- why
- what remains
- risks

Never jump straight to random edits without a plan.

---

# 20. Task Classification Rules

Classify every request into one of these categories before acting:
- architecture
- new feature
- bug fix
- refactor
- migration
- analytics/reporting
- planning/forecasting
- documentation
- investigation

Then choose the correct workflow.

## For architecture
Produce design first.

## For feature work
Produce file-level plan first.

## For migrations
Assess data safety first.

## For forecasting/planning
State assumptions and horizon first.

## For bug fixing
Find root cause before patching.

---

# 21. File and Documentation Responsibilities

Maintain these repo-level docs.

## AGENTS.md
Master operating rules and project cognition.

## docs/architecture/system-overview.md
High-level structure and boundaries.

## docs/architecture/module-dependencies.md
Which modules depend on which.

## docs/domain/
One file per domain area.

## docs/schemas/db-models.md
Tables and relationships.

## docs/schemas/api-contracts.md
API contracts and DTOs.

## docs/schemas/events.md
Published events and meanings.

## docs/adr/
Architecture decision records.

Every material architectural change should update docs.

---

# 22. Coding Standards

## General
- prefer explicit naming
- prefer small focused functions
- avoid hidden side effects
- do not use magic constants without named meaning
- encode business rules in domain/application layers

## API
- use typed request/response models
- validate inputs explicitly
- return stable error structures
- document endpoint purpose

## Database
- use migrations only
- name constraints and indexes clearly
- add created_at / updated_at where applicable
- add foreign keys where valid
- think about query patterns before schema finalization

## Frontend
- TypeScript required
- business logic should stay minimal in components
- forms must validate but backend remains source of truth
- operational screens must optimize speed of use

## Planning/Forecasting Code
- keep formulas traceable
- keep assumptions configurable
- expose reasons for recommendations
- separate deterministic logic from probabilistic logic

---

# 23. Testing Standards

Every important change must include an approach to testing.

## Test layers
- unit tests for domain rules
- service tests for workflows
- API tests for contracts
- migration tests where relevant
- scenario tests for planning engine
- forecast tests for edge scenarios

## Required test scenarios for planning/forecasting
- enough stock and enough capacity
- enough stock but not enough capacity
- enough capacity but material shortage
- route dependency violation
- supplier lead time causing missed due date
- combined demand causing future stockout
- grouping optimization reducing cost/time

---

# 24. Planning Intelligence Guidelines

When implementing planning, calculate at least:
- order required quantity
- material gross quantity
- waste-adjusted quantity
- free stock net of reservation
- machine hours required
- labor hours required
- setup hours required
- earliest feasible start
- estimated finish
- shortage quantity
- shortage due date

When explaining output, always provide reasons:
- blocked due to board shortage
- blocked due to magnet shortage
- delayed due to die-cut machine overload
- launchable after incoming PO on date X

---

# 25. Forecasting Intelligence Guidelines

Forecasting must support at least 3 scenario modes:

## Conservative
Use confirmed orders only.

## Expected
Use confirmed orders + weighted pipeline.

## Aggressive
Use confirmed orders + strong growth assumptions.

Each forecast should output assumptions explicitly.
Do not produce a single unexplained number.

---

# 26. New Packaging Family Expansion Rules

The user may connect another packaging project with a new packaging type.
When this happens:
1. do not break the existing rigid box model
2. model the new family as an extension of PackagingFamily and PackagingSpec
3. add family-specific route templates and BOM patterns
4. reuse shared planning, inventory, costing, and forecasting engines
5. avoid hardcoding special behavior globally if it belongs only to one family

Goal:
One ERP core, many packaging families.

---

# 27. Procurement Logic Rules

Procurement recommendations must consider:
- current free stock
- reserved stock
- incoming confirmed stock
- supplier MOQ
- supplier lead time
- safety stock rules
- forecasted demand
- substitute material options if allowed

Procurement output should include:
- item
- needed quantity
- shortage date
- recommended order date
- preferred supplier
- risk if not purchased

---

# 28. Machine and Work Center Logic

Each machine/work center should support:
- supported operation types
- speed norms
- setup norms
- format constraints
- material constraints
- shift calendar
- maintenance windows
- current queue
- operator requirements
- outsource alternative if any

Never assume one operation equals one machine.
Support alternatives.

---

# 29. Human Labor Logic

Labor planning must support:
- operator skills
- assignment by machine/work center
- effective hours vs nominal hours
- setup work
- run work
- manual assembly work
- overtime possibility flag

Analytics should distinguish:
- effective productive time
- setup time
- waiting time
- downtime
- non-value-added time

---

# 30. Quality Gate Logic

QC must be stage-aware.
Possible checkpoints:
- incoming material inspection
- post-print inspection
- post-lamination inspection
- post-die-cut inspection
- assembly inspection
- final packing inspection

QC may:
- allow progression
- allow with deviation
- block progression
- trigger corrective action

---

# 31. Events and Automation

Prefer clear domain events where useful.
Examples:
- OrderApproved
- SpecApproved
- ProductionOrderReleased
- MaterialReserved
- MaterialShortageDetected
- PurchaseRequestGenerated
- OperationStarted
- OperationCompleted
- QCFailed
- ForecastGenerated
- BottleneckDetected

Use events to decouple modules where appropriate.
Do not over-engineer event complexity if synchronous service calls are simpler.

---

# 32. Suggested MCP/Tool Integrations

If external tools are available, connect these classes of tools:
- database access tool
- file parsing tool for Excel/PDF/specs
- Git/GitHub tool
- ERP data ingestion tool
- cost calculator tool
- document generation tool
- scheduling/optimization tool

Use tools carefully:
- inspect before editing
- explain assumptions before destructive changes
- prefer safe reads first

---

# 33. Default Build Priorities

When unsure, prioritize implementation in this order:
1. domain correctness
2. planning correctness
3. data integrity
4. explainability
5. maintainability
6. UI polish

Do not sacrifice traceability for speed.

---

# 34. Default Implementation Roadmap

If asked what to build first, recommend this order:

## Phase 1
- auth/users/roles
- clients/orders/quotes
- packaging spec core
- materials and warehouse core
- machine/work center core

## Phase 2
- BOM and route templates
- production order release
- MES execution events
- QC basics
- documents/labels

## Phase 3
- costing engine
- capacity planning
- finite scheduling
- shortage detection

## Phase 4
- demand forecast
- material forecast
- procurement forecast
- scenario planning
- analytics dashboards

## Phase 5
- advanced optimization
- AI recommendations
- anomaly detection
- camera/HMI integration

---

# 35. Default Response Format for Development Tasks

Unless user requests otherwise, respond in this structure:
1. objective
2. affected modules
3. plan
4. implementation notes
5. validation
6. risks / next steps

For code changes, always mention files to create or edit.

---

# 36. Instruction for Starting Any New Task

Before working, internally do this checklist:
- Which business capability is this improving?
- Which module owns this behavior?
- Which entities change?
- Which invariants matter?
- Is this deterministic planning or probabilistic forecasting?
- What docs must update?
- What tests prove this?

---

# 37. Instruction for Repo Bootstrap

If the repository is incomplete, bootstrap using this structure:

```text
.cursor/rules/
docs/architecture/
docs/domain/
docs/schemas/
docs/adr/
modules/
packages/
apps/
```

Create missing files as needed:
- AGENTS.md
- docs/architecture/system-overview.md
- docs/architecture/module-dependencies.md
- docs/domain/orders.md
- docs/domain/packaging.md
- docs/domain/planning.md
- docs/domain/forecasting.md
- docs/schemas/db-models.md
- docs/schemas/api-contracts.md
- docs/schemas/events.md

If asked to initialize, generate these first.

---

# 38. Instruction for AGENTS.md Generation

If this file is not yet copied into AGENTS.md, use this file as source material to create a shorter operational AGENTS.md and reference deeper docs from docs/.

AGENTS.md should include:
- mission
- core entities
- invariants
- workflow
- architecture rules
- documentation rules

---

# 39. Instruction for Handling Ambiguity

When information is missing:
- do not collapse the architecture
- make conservative assumptions
- clearly mark assumptions
- prefer extensible models
- do not hardcode customer-specific logic as global truth

---

# 40. Final Standing Orders

You are not here just to write code fragments.
You are here to help build an industrial-grade packaging ERP/MES platform with planning and forecasting intelligence.

Always optimize for:
- correctness
- traceability
- extensibility
- explainability
- production realism

Protect the architecture.
Protect the domain model.
Protect data integrity.
Keep implementation modular.
Keep docs synchronized.
Keep the system ready for new packaging families and future predictive planning.
