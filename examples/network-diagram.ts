import { VisioDocument } from '../src/VisioDocument';
import * as path from 'path';

async function run() {
    console.log("Generating network-topology.vsdx...");
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    // 1. Create Core Switch (Center)
    const coreSwitch = await page.addShape({
        text: "Core Switch\n(L3)",
        x: 4,
        y: 8,
        width: 2,
        height: 1.5,
        fillColor: "#E8E8E8",
        bold: true
    });

    // Fluent Data API Usage
    coreSwitch.addData('ip', { value: '10.0.0.1', label: 'Management IP' })
        .addData('model', { value: 'Cisco Catalyst', label: 'Hardware Model' })
        .addData('vlan', { value: 100, label: 'Mgmt VLAN' })
        .addData('asset_id', { value: 'SW-CORE-99', hidden: true }); // Hidden tracking ID

    // 2. Create Web Server (Left)
    const webServer = await page.addShape({
        text: "Web01\n(Nginx)",
        x: 2,
        y: 5,
        width: 1.5,
        height: 1
    });

    webServer.addData('ip', { value: '10.0.10.5' })
        .addData('role', { value: 'Frontend' })
        .addData('last_patch', { value: new Date() }) // Auto-serialized Date
        .addData('auto_scale', { value: true });      // Boolean

    // 3. Create DB Server (Right)
    const dbServer = await page.addShape({
        text: "DB01\n(Postgres)",
        x: 6,
        y: 5,
        width: 1.5,
        height: 1
    });

    dbServer.addData('ip', { value: '10.0.20.5' })
        .addData('role', { value: 'Database' })
        .addData('is_primary', { value: true, label: 'Primary Node' });

    // 4. Create Firewall (Top)
    const firewall = await page.addShape({
        text: "Edge Firewall",
        x: 4,
        y: 10,
        width: 2,
        height: 1,
        fillColor: "#FFCCCC"
    });
    firewall.addData('policy_ver', { value: '2.5.1' });

    // 5. Connect Components
    // Uplink to Firewall
    await coreSwitch.connectTo(firewall);

    // Servers to Core
    await webServer.connectTo(coreSwitch);
    await dbServer.connectTo(coreSwitch);

    // App Logic connection (Web -> DB)
    await webServer.connectTo(dbServer);

    // 6. Save
    const outPath = path.resolve(__dirname, 'network-topology.vsdx');
    await doc.save(outPath);
    console.log(`Saved to ${outPath}`);
}

run().catch(console.error);
