Based on the current implementation and typical business requirements for diagramming, here are the most significant missing use cases for the Visio SDK:

(Done) v1. Masters and Stencils (Top Priority)
Currently,

addShape
 draws every shape from scratch using raw geometry (LineTo, MoveTo). Real-world Visio usage relies on Masters—reusable templates stored in stencils (like "Network", "Basic Flowchart", "AWS Icons").

Use Case: "Drop a 'Router' icon here" instead of "Draw a box with these dimensions."
Technical Gap: Need to handle visio/masters/master*.xml and link shapes via the Master attribute.

(Done) 2. Multi-Page Support
The current implementation hardcodes page1.xml. Most professional diagrams (like database schemas or floor plans) span multiple tabs.

Use Case: "Create a 'Summary' page and a 'Details' page in the same document."
Technical Gap: Need PageManager to handle visio/pages/page*.xml creation and update visio/pages/pages.xml.

(Done) 3. Shape Data (Custom Properties)
Visio is often used as a "visual database." Shapes usually contain metadata (e.g., a "Server" shape having "IP Address", "Cost", "OS").

Use Case: "Add an 'Employee ID' field to this box that exports to Excel but isn't visible on the drawing."
Technical Gap: Need to verify serialization of <Section N="Property">.

(Done) 4. Images
There is no support for embedding images.

Use Case: "Add a company logo to the header" or "Paste a screenshot of the UI."
Technical Gap: Need to embed binary image data into the ZIP package (visio/media/image1.png) and create a shape that references it using ForeignData.

(Done) 5. Hyperlinks
Shapes often link to other pages or external URLs.

Use Case: "Clicking this 'User' table should open the Jira ticket."
Technical Gap: Implementation of <Section N="Hyperlink">.

(Done) 6. Container Shapes vs. Groups
We implemented "Groups" for tables, but Visio has a special concept called Containers. Containers automatically "grab" shapes dropped into them and resize to fit content.

Use Case: "Drag a new column into the Table, and the Table automatically expands to fit it."
Technical Gap: Specific User Defined Cells and algorithmic behavior in Visio (often harder to fully replicate without the Visio engine, but the XML structures exist).

(Done) 7. Layers
Complex diagrams use layers to toggle visibility (e.g., "Hide all comments").

Use Case: "Create a 'Wireframe' layer and a 'Notes' layer."
Technical Gap: visio/pages/pageN.xml <PageContents><Layers>.
Recommendation: The next most high-value target would likely be Masters/Stencils or Multi-Page Support, as these are fundamental to creating professional-grade diagrams rather than just "drawings."