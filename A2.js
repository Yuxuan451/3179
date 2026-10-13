const TSX = "TSX.xlsx";
const data_gov = "data.gov.xlsx";
const MAP_FILE = "australia-states.json";


// ==========================================
// LOAD Data
// ==========================================

fetch(TSX)
    .then(response => response.arrayBuffer())
    .then(arrayBuffer => {

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        // Get the first worksheet
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        // Convert Excel worksheet into JavaScript objects
        const data = XLSX.utils.sheet_to_json(worksheet);
        // console.log("Excel loaded:");
        // console.log(data);

        const speciesByState = {};
        const stateCentroids = {};
        const statetaxonomicCounts = {};
        const statetstatusCounts = {};
        const dotSpecies = {};
        const dotsByState = {};
        const timeSeriesCounts = {};
        const years = [];
        const sankeyCounts = {};
        const statusYearCounts = {};
        const bumpCounts = {};

        for (let year = 2000; year <= 2021; year++) {
            years.push(String(year));
        }

        data.forEach(row => {
            const state = String(row.State || "").trim();
            const species = String(row.Binomial || "").trim();
            const status = String(row.IUCNStatus || "").trim();
            const family = String(row.Family || "").trim();
            const genus = String(row.Genus || "").trim();
            const order = String(row.Order || "").trim();
            const taxonomic = String(row.TaxonomicGroup || "").trim();
            const lat = Number(row.RegionCentroidLatitude);
            const lon = Number(row.RegionCentroidLongitude);


            // Ignore invalid rows and Commonwealth total
            if (
                !taxonomic ||
                !order ||
                !family ||
                !genus ||
                !species ||
                !state ||
                !status ||
                state === "Commonwealth" ||
                !Number.isFinite(lat) ||
                !Number.isFinite(lon)
            ){
                return
            };

            // species by state
            if (!speciesByState[state]) {
                speciesByState[state] = new Set();
            }
            speciesByState[state].add(species);

            // center location of each state
            if (!stateCentroids[state]) {
                stateCentroids[state] = {
                    latitude: lat,
                    longitude: lon
                };
            }

            // taxonomic group and species
            const taxonomickey = `${state}|${taxonomic}`;
            if (!statetaxonomicCounts[taxonomickey]) {
                statetaxonomicCounts[taxonomickey] = {
                    State: state,
                    taxonomic: taxonomic,
                    species: new Set(),
                    latitude: lat,
                    longitude: lon
                };
            }
            statetaxonomicCounts[taxonomickey].species.add(species);

            // dot species
            const dotkey = `${state}|${species}`;

            if (!dotSpecies[dotkey]) {
                dotSpecies[dotkey] = {
                    State: state,
                    taxonomic: taxonomic,
                    IUCNStatus: status,
                    latitude: lat,
                    longitude: lon
                };
            }

            // status x threat status
            const statuskey = `${state}|${status}`;
            if (!statetstatusCounts[statuskey]) {
                statetstatusCounts[statuskey] = {
                    State: state,
                    IUCNStatus: status,
                    species: new Set(),
                };
            }
            statetstatusCounts[statuskey].species.add(species);

            const baseVal = Number(row["2000"]);

            // Cannot calculate percentage change
            // if species has no 2000 abundance
            if (!Number.isFinite(baseVal) || baseVal === 0) {
                return;
            }
            years.forEach(year => {
                const value = row[year];
                // Keep actual 0 values
                if (value !== undefined &&value !== null &&value !== "") {

                    const abundance = Number(value);
                    if (Number.isFinite(abundance)) {
                        const timekey = `${taxonomic}|${year}`;
                        if (!timeSeriesCounts[timekey]) {
                            timeSeriesCounts[timekey] = {
                                TaxonomicGroup: taxonomic,
                                Year: Number(year),
                                Abundance: 0
                            };
                        }
                        timeSeriesCounts[timekey].Abundance += abundance;
                    }
                }

                const abundance = Number(row[year]);
                if (!Number.isFinite(abundance)) {
                    return;
                }
                const bumpkey = `${state}|${year}`;
                if (!bumpCounts[bumpkey]) {
                    bumpCounts[bumpkey] = {
                        State: state,
                        Year: Number(year),
                        Abundance: 0
                    };
                }
                bumpCounts[bumpkey].Abundance += abundance;
            });

            ["2000", "2021"].forEach(year => {
                const abundance = Number(row[year]);
                if (!Number.isFinite(abundance)) {
                    return;
                }
                const connectDotkey = `${status}|${year}`;
                if (!statusYearCounts[connectDotkey]) {
                    statusYearCounts[connectDotkey] = {
                        Status: status,
                        Year: Number(year),
                        Abundance: 0
                    };
                }
                statusYearCounts[connectDotkey].Abundance += abundance;
            });
            

            // Remove duplicate species
            const sankeykey = `${taxonomic}|${order}|${family}|${species}`;
            if (!sankeyCounts[sankeykey]) {
                sankeyCounts[sankeykey] = {
                    Taxonomic: taxonomic,
                    Order: order,
                    Family: family,
                    Species: species
                };
            }
        });

        const stateCodeMap = {
            "Western Australia": "AU-WA",
            "Northern Territory": "AU-NT",
            "South Australia": "AU-SA",
            "Queensland": "AU-QLD",
            "New South Wales": "AU-NSW",
            "Victoria": "AU-VIC",
            "Tasmania": "AU-TAS",
            "Australian Capital Territory": "AU-ACT"
        };

        const stateCounts = Object.entries(speciesByState).map(([state, speciesSet]) => ({
            State: state,
            state_code: stateCodeMap[state],
            species_count: speciesSet.size
        }));
        // console.log("Species count by state:");
        // console.table(stateCounts);

        const offsets = {
            "Birds": { lat: 0.8, lon: -0.8 },
            "Mammals": { lat: -0.8, lon: 0.8 },
            "Amphibians": { lat: -0.8, lon: -0.8 },
            "Plants": { lat: 0.8, lon: 0.8 }
        };

        
        const proportionalSymbolData = Object.values(statetaxonomicCounts).map(d => {
        
        const offset = offsets[d.taxonomic] || { lat: 0, lon: 0 };
        return {
            
            State: d.State,
            taxonomic: d.taxonomic,
            species_count: d.species.size,
            latitude: d.latitude + offset.lat,
            longitude: d.longitude + offset.lon
        }});
        // console.table(proportionalSymbolData);


        const dotDensityData = Object.values(dotSpecies);
        // console.log("Dot density data:");
        // console.table(dotDensityData);
        dotDensityData.forEach(d => {

            if (!dotsByState[d.State]) {
                dotsByState[d.State] = [];
            }

            dotsByState[d.State].push(d);
        });
        Object.values(dotsByState).forEach(stateDots => {

            stateDots.forEach((d, i) => {

                const angle = i * 137.5;
                const radius = 0.15 + ((i % 8) * 0.04);
                const angleRad = angle * Math.PI / 180;

                d.longitude += Math.cos(angleRad) * radius;
                d.latitude += Math.sin(angleRad) * radius;
            });
        });

        const timeSeriesData = Object.values(timeSeriesCounts);
        const lineChartData = [];
        const BASE_YEAR = 2000;
        // Store abundance in 2000 for each taxonomic group
        const baselineValues = {};

        timeSeriesData.forEach(d => {
            if (d.Year === BASE_YEAR) {
                baselineValues[d.TaxonomicGroup] = d.Abundance;
            }

        });


        // Calculate percentage change from 2000
        timeSeriesData.forEach(d => {

            const baseVal = baselineValues[d.TaxonomicGroup];

            // Ignore groups with no valid baseline
            if (!baseVal || baseVal === 0) {
                return;
            }
            console.log("Baseline values:");
            console.table(baselineValues);

            console.log("All time series data:");
            console.table(timeSeriesData);

            const pctChange =
                ((d.Abundance - baseVal) / baseVal);

            lineChartData.push({
                Year: d.Year,
                TaxonomicGroup: d.TaxonomicGroup,
                PctChange: pctChange
            });

        });

        console.log("Percentage change data:");
        console.table(lineChartData);

        // SANKEYDIAGRAM
        const sankeyData = Object.values(sankeyCounts);
        console.log("Sankey data:");
        console.table(sankeyData);
        // const sankeyLinks = [];
        const linkCounts = {};

        sankeyData.forEach(d => {

            // Taxonomic Group → Family
            const groupFamily =`${d.Taxonomic}|${d.Order}`;
            if (!linkCounts[groupFamily]) {
                linkCounts[groupFamily] = {
                    source: d.Taxonomic,
                    target: d.Order,
                    value: 0
                };
            }
            linkCounts[groupFamily].value++;

            // Family → Genus
            const familyGenus =`${d.Order}|${d.Family}`;
            if (!linkCounts[familyGenus]) {
                linkCounts[familyGenus] = {
                    source: d.Order,
                    target: d.Family,
                    value: 0
                };
            }
            linkCounts[familyGenus].value++;

        });
        const sankeyLinks = Object.values(linkCounts);
        console.table(sankeyLinks);

        // bump chart
        const stateYearData = Object.values(bumpCounts);
        const bumpChartData = [];
        years.forEach(year => {
            const yearData = stateYearData
                .filter(d => d.Year === Number(year))
                .sort((a, b) => b.Abundance - a.Abundance);

            yearData.forEach((d, index) => {
                bumpChartData.push({
                    State: d.State,
                    Year: d.Year,
                    Abundance: d.Abundance,
                    Rank: index + 1
                });
            });

        });
        console.table(bumpChartData);
        createBumpChart(bumpChartData);

        //connected dot plot
        const connectedDotData = Object.values(statusYearCounts);
        const statusBaselines = {};

        connectedDotData.forEach(d => {

            if (d.Year === 2000) {
                statusBaselines[d.Status] = d.Abundance;
            }

        });
        const connectedDotPctData = connectedDotData
            .filter(d => {

                const baseline = statusBaselines[d.Status];

                return baseline !== undefined &&
                    baseline !== 0;

            })
            .map(d => {

                const baseline = statusBaselines[d.Status];

                return {
                    Status: d.Status,
                    Year: d.Year,
                    Abundance: d.Abundance,
                    PctChange:
                        ((d.Abundance - baseline) / baseline) * 100
                };

            });

        console.table(connectedDotPctData);

        // LOAD MAP
        fetch(MAP_FILE)
            .then(response => response.json())
            .then(mapData => {
                createMap(stateCounts, proportionalSymbolData, dotDensityData, mapData);
                createBumpChart(bumpChartData);
                createTimeGraph(lineChartData, connectedDotPctData);
                createSankey(sankeyData);
                const taxonomicFilter = document.getElementById("taxonomicFilter");
                taxonomicFilter.addEventListener("change", function () {

                    const selectedGroup = this.value;

                    console.log(
                        "Selected taxonomic group:",
                        selectedGroup
                    );

                    createSankey(
                        sankeyData,
                        selectedGroup
                    );
                });
            })
            .catch(error => {
                console.error("Error loading australia-states.json:", error);
            });
        })
        .catch(error => {
            console.error("Error loading CSV:", error);
        });

// DATA.GOV DATASET
fetch(data_gov)
    .then(response => response.arrayBuffer())
    .then(arrayBuffer => {

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        // Get the first worksheet
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        // Convert Excel worksheet into JavaScript objects
        const data = XLSX.utils.sheet_to_json(worksheet);
        // console.log("Excel loaded:");
        // console.log(data);

        // const speciesByStateStatus = {};
        const stateColumns = ["ACT","NSW", "NT", "QLD", "SA","TAS","VIC","WA"];
        const kingdomCounts = {};
        const radialCounts = {};
        const sunburstNodes = new Map();

        function addSunburstNode(id, parent, name, level) {
            if (!sunburstNodes.has(id)) {
                sunburstNodes.set(id, {
                    id: id,
                    parent: parent,
                    name: name,
                    level: level,
                    species: new Set()
                });
            }

            return sunburstNodes.get(id);
        }

        data.forEach(row => {
            const status = String(row["Threatened status"] || "").trim();
            const species = String(row["Scientific Name"] || "").trim();
            const kingdom = String(row["Kingdom"] || "").trim();
            const className = String(row["Class"] || "").trim();
            const family = String(row["Family"] || "").trim();
            const genus = String(row["Genus"] || "").trim();
            // Ignore incomplete rows
            if (!species || !status || !kingdom || !className ||!family ||!genus) {
                return;
            }

            let group;
            // group animal and platns together
            if (kingdom === "Animalia") {
                group = "Animals";
            } else if (kingdom === "Plantae") {
                group = "Plants";
            } else {
                return;
            }


            // Check each Australian state
            stateColumns.forEach(state => {
                const value = String(row[state] || "").trim();
                // Only count if species occurs in state
                if (value.toLowerCase() === "yes") {

                    // const key =`${state}|${status}`;
                    // // state + status combination
                    // if (!speciesByStateStatus[key]) {
                    //     speciesByStateStatus[key] = {
                    //         State: state,
                    //         ThreatStatus: status,
                    //         species: new Set()
                    //     };
                    // }
                    // // Add unique species
                    // speciesByStateStatus[key].species.add(species);

                    const kingdomkey = `${state}|${group}`;
                    if (!kingdomCounts[kingdomkey]) {
                                kingdomCounts[kingdomkey] = {
                                    State: state,
                                    group: group,
                                    species: new Set()
                                };
                            }
                    // Add unique species
                    kingdomCounts[kingdomkey].species.add(species);

                    // radial stacked bar 
                    if (!radialCounts[state]) {
                        radialCounts[state] = {};
                    }

                    if (!radialCounts[state][status]) {
                        radialCounts[state][status] = new Set();
                    }

                    radialCounts[state][status].add(species);
                            }
                        });
            // ROOT
            const root = addSunburstNode(
                "root",
                null,
                "Australia",
                "Root"
            );

            // KINGDOM
            const kingdomId = `kingdom|${kingdom}`;

            const kingdomNode = addSunburstNode(
                kingdomId,
                "root",
                kingdom,
                "Kingdom"
            );

            // CLASS
            const classId = `class|${kingdom}|${className}`;

            const classNode = addSunburstNode(
                classId,
                kingdomId,
                className,
                "Class"
            );

            // FAMILY
            const familyId = `family|${kingdom}|${className}|${family}`;

            const familyNode = addSunburstNode(
                familyId,
                classId,
                family,
                "Family"
            );

            // GENUS
            const genusId =
                `genus|${kingdom}|${className}|${family}|${genus}`;

            const genusNode = addSunburstNode(
                genusId,
                familyId,
                genus,
                "Genus"
            );

            // Add species to EVERY ancestor
            root.species.add(species);
            kingdomNode.species.add(species);
            classNode.species.add(species);
            familyNode.species.add(species);
            genusNode.species.add(species);

    
        });

        // const heatmapData = Object.values(speciesByStateStatus).map(d => {
        // return {
        //     State: d.State,
        //     ThreatStatus: d.ThreatStatus,
        //     species_count: d.species.size
        // }});
        // // console.table(heatmapData);

        const tornadochartData = Object.values(kingdomCounts).map(d => ({
            State: d.State,
            group: d.group,

            // Animals go to the left
            // Plants go to the right
            species_count:
                d.group === "Animals"
                    ? -d.species.size
                    : d.species.size
        }));
        // console.table(tornadochartData);

        const radialStatusOrder = [
            "Critically Endangered",
            "Endangered",
            "Vulnerable",
            "Conservation Dependent",
            "Extinct",
            "Extinct in the wild"
        ];
        const radialStackedData = [];
        stateColumns.forEach(state => {
            let cumulative = 0;
            radialStatusOrder.forEach(status => {
                const count =
                    radialCounts[state]?.[status]?.size || 0;
                if (count === 0) return;
                const startValue = cumulative;
                const endValue = cumulative + count;
                radialStackedData.push({
                    State: state,
                    ThreatStatus: status,
                    species_count: count,
                    startValue: startValue,
                    endValue: endValue
                });
                cumulative = endValue;
            });
        });

        // console.log("Radial stacked data:");
        // console.table(radialStackedData);

        const hierarchyNodes = Array.from(sunburstNodes.values()).map(d => ({
            id: d.id,
            parent: d.parent,
            name: d.name,
            level: d.level,
            species_count: d.species.size
        }));

        // const hierarchyNodes = Array.from(sunburstNodes.values()).map(d => ({
        //     id: d.id,
        //     parent: d.parent,
        //     name: d.name,
        //     level: d.level,

        //     species_count:
        //         d.level === "Genus"
        //             ? d.species.size
        //             : 0
        // }));

        console.log("SUNBURST DATA:");
        console.table(hierarchyNodes);
        console.log("Number of nodes:", hierarchyNodes.length);

        // LOAD MAP
        fetch(MAP_FILE)
            .then(response => response.json())
            .then(mapData => {
                // createheatMap(heatmapData, mapData);
                createChart(tornadochartData, radialStackedData, hierarchyNodes)
            })
            .catch(error => {
                console.error("Error loading australia-states.json:", error);
            });
        })
        .catch(error => {
            console.error("Error loading CSV:", error);
        });

// ==========================================
// CREATE MAP FUNCTION
// ==========================================
function createMap(stateCounts, proportionalSymbolData, dotDensityData, mapData) {
    fetch("Maps/choropleth_map.json")
        .then(response => response.json())
        .then(spec => {
            // MAP DATA
            spec.layer[0].data = {
                values: mapData,
                format: {
                    type: "topojson",
                    feature: "ne_10m_admin_1_states_provinces"
                }
            };
            // STATE COUNT DATA
            spec.layer[0].transform[1].from = {
                data: {values: stateCounts},
                key: "state_code",
                fields: [
                    "State",
                    "species_count"
                ]
            };
            vegaEmbed("#choropleth_map", spec, {
                actions: false
            })
            .then(result => {console.log("Map 1 successfully displayed.");})
            .catch(error => {
                console.error("Vega-Lite error:");
                console.error(error);
            });
        })
        .catch(error => {
            console.error("Error loading choropleth_map.json:", error);
        });

    fetch("Maps/proportional_symbol_map.json")
        .then(response => response.json())
        .then(spec => {
            // Insert our calculated state data
            spec.layer[1].data.values = proportionalSymbolData;
            vegaEmbed("#proportional_symbol_map", spec, {
                actions: false
            })
            .then(() => {
                console.log("Proportional Symbol Map successfully displayed.");
            })
            .catch(error => {console.error("Vega-Lite error:", error);});
        })
        .catch(error => {
            console.error("Error loading proportional_symbol_map.json:", error);
        });

    fetch("Maps/dot_density_map.json")
        .then(response => response.json())
        .then(spec => {
            spec.layer[1].data.values = dotDensityData;
            vegaEmbed("#dot_density_map", spec,
                {
                    actions: false
                }
            )
            .then(() => {
                console.log("Dot Density Map successfully displayed.");
            })
            .catch(error => {
                console.error("Vega-Lite error:", error);
            });
        })
        .catch(error => {
            console.error("Error loading dot_density_map.json:",error);
        });
}

// HEATMAP
// function createheatMap(heatmapData, mapData) {
//     fetch("heatmap.json")
//     .then(response => response.json())
//     .then(spec => {

//         spec.data.values = heatmapData;
//         vegaEmbed("#heatmap", spec,{actions: false})
//         .then(() => {
//             console.log("Heatmap successfully displayed.");
//         })
//         .catch(error => {
//             console.error("Vega-Lite error:");
//             console.error(error);
//         });
//     })
//     .catch(error => {
//         console.error("Error loading heatmap.json:", error);
//     });

// }

function createChart(tornadochartData, radialStackedData, hierarchyNodes) {

    fetch("Charts/tornado_chart.json")
        .then(response => response.json())
        .then(spec => {
            spec.data.values = tornadochartData;
            vegaEmbed("#tornado_chart",spec,{actions: false});
        })
        .catch(error => {
            console.error("tornado chart error:", error);
        });

    fetch("Charts/radial_stacked_bar.json")
        .then(response => response.json())
        .then(spec => {

            spec.data[0].values = radialStackedData;

            vegaEmbed("#radial_stacked_bar", spec, {
                actions: false,
                renderer: "svg"
            })
            .then(() => {
                console.log("Radial stacked bar successfully displayed.");
            })
            .catch(error => {
                console.error("Radial stacked bar Vega error:", error);
            });

        })
        .catch(error => {
            console.error(
                "Error loading radial_stacked_bar.json:",
                error
            );
        });

    fetch("Charts/sunburst.json")
            .then(response => response.json())
            .then(spec => {

                spec.data[0].values = hierarchyNodes;

                vegaEmbed("#sunburst", spec, {
                    actions: false,
                    renderer: "svg"
                })
                .then(() => {
                    console.log("Sunburst successfully displayed.");
                })
                .catch(error => {
                    console.error(
                        "Sunburst Vega error:",
                        error
                    );
                });

            })
            .catch(error => {
                console.error(
                    "Error loading sunburst.json:",
                    error
                );
            });
}

function createTimeGraph(lineChartData, connectedDotPctData) {
    fetch("Time-Series/stream_graph.json")
        .then(response => response.json())
        .then(spec => {
            // Insert aggregated time-series data
            spec.data.values = lineChartData;
            vegaEmbed("#stream_graph",spec,{ actions: false });
        })
        .catch(error => {
            console.error("Time series error:", error);});

    fetch("Time-Series/connected_dot.json")
        .then(response => response.json())
        .then(spec => {

            spec.data.values = connectedDotPctData;

            vegaEmbed("#connected_dot", spec, {
                actions: false,
                renderer: "svg"
            })
            .then(() => {
                console.log("Connected dot plot successfully displayed.");
            })
            .catch(error => {
                console.error(
                    "Connected dot plot Vega error:",
                    error
                );
            });

        })
        .catch(error => {
            console.error(
                "Error loading connected_dot_plot.json:",
                error
            );
        });
}

function createBumpChart(bumpChartData) {

    fetch("Time-Series/bump_chart.json")
        .then(response => response.json())
        .then(spec => {

            // ==========================================
            // INSERT BUMP CHART DATA
            // ==========================================

            if (!spec.data) {
                spec.data = {};
            }

            spec.data.values = bumpChartData;


            // ==========================================
            // RENDER
            // ==========================================

            vegaEmbed(
                "#bump_chart",
                spec,
                {
                    actions: false,
                    renderer: "svg"
                }
            )
            .then(() => {
                console.log(
                    "Bump chart successfully displayed."
                );
            })
            .catch(error => {
                console.error(
                    "Bump chart Vega error:",
                    error
                );
            });

        })
        .catch(error => {
            console.error(
                "Error loading bump_chart.json:",
                error
            );
        });
}

// ==========================================
// SANKEY GLOBAL VARIABLES
// ==========================================

let rawSankeyData = [];
let sankeyView = null;


// ==========================================
// CREATE SANKEY
// ==========================================
// This function only prepares the data.
// The actual filtering and visualisation
// is handled by updateSankey().
// ==========================================

function createSankey(sankeyData) {

    // Store the complete Sankey dataset
    rawSankeyData = sankeyData;

    console.log("Sankey raw data:");
    console.table(rawSankeyData);


    // ==========================================
    // INITIAL DISPLAY
    // ==========================================

    updateSankey("All");


    // ==========================================
    // TAXONOMIC GROUP FILTER
    // ==========================================

    const filter =
        document.getElementById("taxonomicFilter");

    if (filter) {

        filter.addEventListener("change", function () {

            const selectedGroup = this.value;

            console.log(
                "Sankey selected group:",
                selectedGroup
            );

            updateSankey(selectedGroup);

        });

    } else {

        console.warn(
            "taxonomicFilter element was not found."
        );

    }
}


// ==========================================
// UPDATE SANKEY
// ==========================================
// Filters the data, creates nodes/links,
// calculates the D3 Sankey layout and
// renders the result in Vega.
// ==========================================


async function updateSankey(selectedGroup = "All") {


    // ==========================================
    // 1. FILTER DATA
    // ==========================================

    const filteredData =
        selectedGroup === "All"
            ? rawSankeyData
            : rawSankeyData.filter(d =>
                String(d.Taxonomic || "").trim()
                === selectedGroup
            );


    console.log(
        "Selected Sankey group:",
        selectedGroup
    );

    console.log(
        "Filtered Sankey rows:",
        filteredData.length
    );

    console.table(filteredData);


    // ==========================================
    // 2. CREATE NODES
    // ==========================================

    const nodes = [];
    const nodeMap = new Map();


    function addNode(name, level) {

        name = String(name || "").trim();

        if (!name) {
            return null;
        }


        const id = `${level}|${name}`;


        if (!nodeMap.has(id)) {

            const node = {

                id: id,

                name: name,

                level: level

            };


            nodeMap.set(id, node);

            nodes.push(node);

        }


        return id;
    }


    // ==========================================
    // 3. CREATE LINKS
    // ==========================================

    const linkCounts = new Map();


    filteredData.forEach(row => {

        const group =
            String(row.Taxonomic || "").trim();

        const order =
            String(row.Order || "").trim();

        const family =
            String(row.Family || "").trim();


        // Ignore incomplete rows

        if (
            !group ||
            !order ||
            !family
        ) {
            return;
        }


        // --------------------------------------
        // CREATE NODES
        // --------------------------------------

        const groupId =
            addNode(group, "group");

        const orderId =
            addNode(order, "order");

        const familyId =
            addNode(family, "family");


        // --------------------------------------
        // GROUP → ORDER
        // --------------------------------------

        const groupOrderKey =
            `${groupId}->${orderId}`;


        if (!linkCounts.has(groupOrderKey)) {

            linkCounts.set(
                groupOrderKey,
                {

                    source: groupId,

                    target: orderId,

                    value: 0

                }
            );

        }


        linkCounts
            .get(groupOrderKey)
            .value++;


        // --------------------------------------
        // ORDER → FAMILY
        // --------------------------------------

        const orderFamilyKey =
            `${orderId}->${familyId}`;


        if (!linkCounts.has(orderFamilyKey)) {

            linkCounts.set(
                orderFamilyKey,
                {

                    source: orderId,

                    target: familyId,

                    value: 0

                }
            );

        }


        linkCounts
            .get(orderFamilyKey)
            .value++;

    });


    const links =
        Array.from(linkCounts.values());


    console.log(
        "Sankey nodes before layout:",
        nodes
    );

    console.log(
        "Sankey links before layout:",
        links
    );


    // ==========================================
    // 4. HANDLE EMPTY DATA
    // ==========================================

    if (nodes.length === 0) {

        console.warn(
            "No Sankey data for:",
            selectedGroup
        );

        return;

    }


    // ==========================================
    // 5. COUNT NODES PER COLUMN
    // ==========================================

    const groupCount =
        nodes.filter(
            d => d.level === "group"
        ).length;

    const orderCount =
        nodes.filter(
            d => d.level === "order"
        ).length;

    const familyCount =
        nodes.filter(
            d => d.level === "family"
        ).length;


    const maxNodes =
        Math.max(
            groupCount,
            orderCount,
            familyCount
        );


    // ==========================================
    // 6. DYNAMIC HEIGHT
    // ==========================================

    const topMargin = 30;
    const bottomMargin = 40;

    // Give each node enough vertical space.
    // The height can grow for dense groups instead
    // of squeezing every group into 700px.
    const nodeHeightTarget = 28;

    const dynamicHeight =
        Math.min(
            500,
            Math.min(
                1800,
                maxNodes * nodeHeightTarget + topMargin + bottomMargin
            )
        );


    // ==========================================
    // 7. CHART SIZE
    // ==========================================

    const chartWidth = 850;

    // Actual Sankey width
    // Smaller than entire chart so
    // there is equal whitespace around it.

// Keep Sankey centered and compact
    const sideMargin = 100;

    const sankeyLeft = sideMargin;
    const sankeyRight = chartWidth - sideMargin;


// ==========================================
// ADJUST NODE PADDING
// ==========================================

    // More nodes = slightly smaller spacing
    let nodePadding = 14;

    if (maxNodes > 15) {
        nodePadding = 10;
    }

    if (maxNodes > 25) {
        nodePadding = 7;
    }

    if (maxNodes > 35) {
        nodePadding = 5;
    }


    // ==========================================
    // 8. D3 SANKEY LAYOUT
    // ==========================================
    const sankeyLayout = d3.sankey()
        .nodeId(d => d.id)
        .nodeWidth(16)
        .nodePadding(nodePadding)
        .nodeAlign(d3.sankeyLeft)
        .extent([
            [sankeyLeft, topMargin],
            [sankeyRight, dynamicHeight - bottomMargin]
        ]);


    // ==========================================
    // 9. RUN SANKEY LAYOUT
    // ==========================================

    const graph =
        sankeyLayout({

            nodes:
                nodes.map(d => ({
                    ...d
                })),

            links:
                links.map(d => ({
                    ...d
                }))

        });


    // ==========================================
    // 10. PREPARE VEGA NODES
    // ==========================================

    const vegaNodes =
        graph.nodes.map(node => ({

            id: node.id,

            name: node.name,

            level: node.level,

            x0: node.x0,

            x1: node.x1,

            y0: node.y0,

            y1: node.y1,

            nodeHeight:
                node.y1 - node.y0

        }));


    // ==========================================
    // 11. PREPARE VEGA LINKS
    // ==========================================

    const vegaLinks =
        graph.links.map(link => ({

            source:
                link.source.id,

            target:
                link.target.id,

            value:
                link.value,

            x0:
                link.source.x1,

            x1:
                link.target.x0,

            y0:
                link.y0,

            y1:
                link.y1,

            width:
                Math.max(
                    1,
                    link.width
                ),

            path:
                d3.sankeyLinkHorizontal()(link)

        }));


    console.log(
        "Final Sankey nodes:",
        vegaNodes
    );

    console.log(
        "Final Sankey links:",
        vegaLinks
    );


    // ==========================================
    // 12. LOAD VEGA JSON
    // ==========================================

    try {

        const response =
            await fetch(
                "Charts/sankey_diagram.json"
            );


        if (!response.ok) {

            throw new Error(
                `Failed to load sankey_diagram.json: ${response.status}`
            );

        }


        const spec =
            await response.json();


        // ======================================
        // 13. SET SIZE
        // ======================================

        spec.width =
            chartWidth;

        spec.height =
            dynamicHeight;


        // ======================================
        // 14. FIND DATASETS
        // ======================================

        const nodeData =
            spec.data.find(
                d => d.name === "nodes"
            );

        const linkData =
            spec.data.find(
                d => d.name === "links"
            );


        if (!nodeData || !linkData) {

            console.error(
                "Could not find nodes or links in sankey_diagram.json"
            );

            return;

        }


        // ======================================
        // 15. INSERT DATA
        // ======================================

        nodeData.values =
            vegaNodes;

        linkData.values =
            vegaLinks;


        // ======================================
        // 16. UPDATE TITLE
        // ======================================

        spec.title.text =
            selectedGroup === "All"

                ? "Taxonomic Hierarchy of Threatened Species"

                : `${selectedGroup} Taxonomic Hierarchy`;


        // ======================================
        // 17. RENDER
        // ======================================

        await vegaEmbed(

            "#sankey_diagram",

            spec,

            {

                actions: false,

                renderer: "svg"

            }

        );


        console.log(
            "Sankey successfully displayed:",
            selectedGroup
        );


    } catch (error) {

        console.error(
            "Sankey rendering error:",
            error
        );

    }

}

