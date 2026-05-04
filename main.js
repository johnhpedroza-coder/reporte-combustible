geotab.addin.fuelEfficiencyReport = (api, state) => {
    let dataCache = {};

    const init = () => {
        const to = new Date();
        const from = new Date();
        from.setMonth(to.getMonth() - 1);
        document.getElementById('dateFrom').value = from.toISOString().split('T')[0];
        document.getElementById('dateTo').value = to.toISOString().split('T')[0];
    };

    const loadData = async () => {
        const from = new Date(document.getElementById('dateFrom').value).toISOString();
        const to = new Date(document.getElementById('dateTo').value).toISOString();
        
        const [trips, devices, users] = await Promise.all([
            api.call("Get", { typeName: "Trip", search: { fromDate: from, toDate: to } }),
            api.call("Get", { typeName: "Device" }),
            api.call("Get", { typeName: "User" })
        ]);

        const devMap = Object.fromEntries(devices.map(d => [d.id, d.name]));
        const userMap = Object.fromEntries(users.map(u => [u.id, `${u.firstName} ${u.lastName}`]));
        const grouped = {};

        trips.forEach(t => {
            const dId = t.device.id;
            const uId = t.driver.id;
            if (!grouped[dId]) grouped[dId] = { name: devMap[dId] || dId, drivers: {} };
            if (!grouped[dId].drivers[uId]) grouped[dId].drivers[uId] = { name: userMap[uId] || "Anon", dist: 0, fuel: 0 };
            grouped[dId].drivers[uId].dist += (t.distance || 0);
            grouped[dId].drivers[uId].fuel += (t.fuelUsed || 0);
        });

        render(grouped);
        dataCache = grouped;
    };

    const render = (data) => {
        const body = document.getElementById('tableBody');
        const chart = document.getElementById('chart');
        body.innerHTML = ""; chart.innerHTML = "";

        for (const dId in data) {
            body.innerHTML += `<tr class="vehicle-row"><td colspan="4">🚚 ${data[dId].name}</td></tr>`;
            for (const uId in data[dId].drivers) {
                const drv = data[dId].drivers[uId];
                const eff = drv.fuel > 0 ? (drv.dist / drv.fuel) : 0;
                body.innerHTML += `<tr><td>👤 ${drv.name}</td><td>${drv.dist.toFixed(1)}</td><td>${drv.fuel.toFixed(1)}</td><td>${eff.toFixed(2)}</td></tr>`;
                
                const bar = document.createElement('div');
                bar.className = 'bar';
                bar.style.height = Math.min(eff * 5, 100) + "%";
                bar.title = `${drv.name}: ${eff.toFixed(2)}`;
                chart.appendChild(bar);
            }
        }
    };

    return {
        initialize(api, state, callback) {
            init();
            document.getElementById('btnRun').addEventListener('click', loadData);
            document.getElementById('btnExcel').addEventListener('click', () => {
                const rows = [["Vehiculo", "Conductor", "KM", "L", "km/L"]];
                for(let v in dataCache) for(let u in dataCache[v].drivers) {
                    const d = dataCache[v].drivers[u];
                    rows.push([dataCache[v].name, d.name, d.dist, d.fuel, d.fuel > 0 ? d.dist/d.fuel : 0]);
                }
                const ws = XLSX.utils.aoa_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Data");
                XLSX.writeFile(wb, "Eficiencia.xlsx");
            });
            callback();
        },
        focus(api, state) { loadData(); }
    };
};
