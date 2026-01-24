import { VisioDocument } from '../src/index';
// @ts-ignore
import fs from 'fs';

async function run() {
    const doc = await VisioDocument.create();

    // Page 1: Dashboard
    const dashboard = doc.pages[0];

    // Create Detail Page
    const detailsPage = await doc.addPage('Server Details');

    console.log('Creating Dashboard...');

    // 1. External Link (JIRA)
    const jiraBox = await dashboard.addShape({
        text: 'View Ticket (JIRA)',
        x: 2, y: 8,
        width: 3, height: 1,
        fillColor: '#DEEBFF' // Blue
    });

    await jiraBox.toUrl('https://jira.atlassian.com', 'Open JIRA Ticket');

    // 2. Internal Link (Drill Down)
    const drillDownBox = await dashboard.addShape({
        text: 'Drill Down to Details',
        x: 6, y: 8,
        width: 3, height: 1,
        fillColor: '#E3FCEF' // Green
    });

    await drillDownBox.toPage(detailsPage, 'Go to Details');

    // 3. Complex URL (Search)
    const googleBox = await dashboard.addShape({
        text: 'Search Error',
        x: 4, y: 6,
        width: 2, height: 1
    });

    await googleBox.toUrl('https://google.com/search?q=visio+xml+error&t=1', 'Search Google');

    // Details Page Content
    await detailsPage.addShape({
        text: 'Server Logs (Details)',
        x: 4, y: 8,
        width: 4, height: 2
    });

    const backBtn = await detailsPage.addShape({
        text: 'Back to Dashboard',
        x: 1, y: 10,
        width: 2, height: 0.5
    });

    // Link back to Dashboard (Page-1 is usually named 'Page-1' by default unless renamed)
    // Actually, doc.pages[0].name is likely 'Page-1'.
    await backBtn.toPage(dashboard, 'Go Back');

    // Save
    const buffer = await doc.save();
    fs.writeFileSync('examples/hyperlinks_demo.vsdx', buffer);
    console.log('Generated examples/hyperlinks_demo.vsdx');
}

run().catch(console.error);
