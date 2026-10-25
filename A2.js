const TSX = "TSX.xlsx";
const data_gov = "data.gov.xlsx";
const MAP_FILE = "australia-states.json";
// ==========================================
// LOAD Data
// ==========================================

fetch(TSX)
    .then(response => response.arrayBuffer())
    .then(arrayBuffer => {
        const workbook = XLSX.read(arrayBuffer, {type: "array"});
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet);
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
        const statusRanks = {
            "Least Concern": 1,
            "Near Threatened": 2,
            "Conservation Dependent": 3,
            "Vulnerable": 4,
            "Endangered": 5,
            "Critically Endangered": 6
        };
        const speciesMap = new Map();
        const speciesByRegion = new Map();

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
            const region = String(row.Region || "").trim();
            const taxonomic = String(row.TaxonomicGroup || "").trim();
            const lat = Number(row.RegionCentroidLatitude);
            const lon = Number(row.RegionCentroidLongitude);
       

            if (!taxonomic ||!order ||!family ||!genus ||!species ||!state ||!status ||state === "Commonwealth" ||!Number.isFinite(lat) ||!Number.isFinite(lon)
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
            const dotkey = species;
            if (!dotSpecies[dotkey]) {
                dotSpecies[dotkey] = {
                    Binomial: species,
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
                // bump chart
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
            
            // Sankey Diagram: Remove duplicate species
            const sankeykey = `${taxonomic}|${order}|${family}|${species}`;
            if (!sankeyCounts[sankeykey]) {
                sankeyCounts[sankeykey] = {
                    Taxonomic: taxonomic,
                    Order: order,
                    Family: family,
                    Species: species
                };
            }
            if (!speciesMap.has(species)) {
                speciesMap.set(species, {
                    Binomial: species,
                    TaxonomicGroup: taxonomic,
                    MaxStatus: status,
                    States: new Set(),
                    Regions: new Set()
                });
            }

            // TREEMAP
            const treekey = `${state}|${region}`;
            if (!speciesByRegion.has(treekey)) {
                speciesByRegion.set(treekey, {
                    state: state,
                    region: region,
                    species: new Set()
                });
            }
            speciesByRegion.get(treekey).species.add(species);
        });

    // ===============================================================================
    //CHOROPLETH MAP
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
    const minCount = Math.min(...stateCounts.map(d => d.species_count));
    const maxCount = Math.max(...stateCounts.map(d => d.species_count));
    stateCounts.forEach(d => {
        d.standardised_count =maxCount === minCount? 0: (d.species_count - minCount) /(maxCount - minCount);
    });

    // ===============================================================================
    // PROPORTIONAL SYMBOL MAP
    const proportionalSymbolData = Object.values(statetaxonomicCounts).map(d => {
    const offsets = {
        "Birds": { lat: 0.8, lon: -0.8 },
        "Mammals": { lat: -0.8, lon: 0.8 },
        "Amphibians": { lat: -0.8, lon: -0.8 },
        "Plants": { lat: 0.8, lon: 0.8 }
    };
    const offset = offsets[d.taxonomic] || { lat: 0, lon: 0 };
    return {
        State: d.State,
        taxonomic: d.taxonomic,
        species_count: d.species.size,
        latitude: d.latitude + offset.lat,
        longitude: d.longitude + offset.lon
    }});

    // ===============================================================================
    // DOT DENSITY MAP
    const dotDensityData = Object.values(dotSpecies);
    const dotsByCoordinate = {};

    dotDensityData.forEach(d => {
        const coordinateKey = `${d.longitude}|${d.latitude}`;
        if (!dotsByCoordinate[coordinateKey]) {
            dotsByCoordinate[coordinateKey] = [];
        }
        dotsByCoordinate[coordinateKey].push(d);
    });

    // Count unique species at each location, then apply jitter
    Object.values(dotsByCoordinate).forEach(locationDots => {
        const uniqueSpecies = new Set(
            locationDots.map(d => d.Binomial)
        );
        locationDots.forEach((d, i) => {
            d.species_count = uniqueSpecies.size;
            if (locationDots.length > 1) {
                const angle = i * 137.508 * Math.PI / 180;
                const radius = 0.012 * Math.sqrt(i + 1);
                d.displayLongitude =d.longitude + Math.cos(angle) * radius - 0.3;
                d.displayLatitude =d.latitude + Math.sin(angle) * radius;
            } else {
                d.displayLongitude = d.longitude - 0.3;
                d.displayLatitude = d.latitude;
            }
        });
    });

    // ===============================================================================
    // MULTI-LINE GRAPH
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
        if (!baseVal || baseVal === 0) {
            return;
        }
        const pctChange =((d.Abundance - baseVal) / baseVal);
        lineChartData.push({
            Year: d.Year,
            TaxonomicGroup: d.TaxonomicGroup,
            PctChange: pctChange
        });

    });

    // ===============================================================================
    // SANKEY DIAGRAM
    const sankeyData = Object.values(sankeyCounts);
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

    // ===============================================================================
    // BUMP CHART
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
    createBumpChart(bumpChartData);

    // ===============================================================================
    //CONNECTED DOT PLOT
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

    // ===============================================================================
    // TREEMAP
    const regionCounts = [...speciesByRegion.values()].map(d => ({
        state: d.state,
        region: d.region,
        count: d.species.size
        }))
        .filter(d => d.count > 0);
    const rows = [
        {
            id: "Australia",
            parent: null,
            value: 0,
            name: "Australia",
            depth: 0
        }
    ];
    const stateGroups = new Map();
    
    regionCounts.forEach(d => {
        if (!stateGroups.has(d.state)) {
            stateGroups.set(d.state, []);
        }

        stateGroups.get(d.state).push(d);
    });
    for (const [state, regions] of stateGroups) {
        const stateId = `state:${state}`;
        rows.push({
            id: stateId,
            parent: "Australia",
            value: regions.reduce((sum, d) => sum + d.count, 0),
            name: state,
            depth: 1
        });

        regions
            .sort((a, b) => b.count - a.count)
            .forEach(d => {
                rows.push({
                    id: `region:${state}:${d.region}`,
                    parent: stateId,
                    value: d.count,
                    name: d.region,
                    state: d.state,
                    region: d.region,
                    depth: 2
                });
            });
    };
    console.table(rows)

        // ===============================================================================
        // LOAD DATA
        fetch(MAP_FILE)
            .then(response => response.json())
            .then(mapData => {
                createMap(stateCounts, proportionalSymbolData, dotDensityData, mapData);
                createBumpChart(bumpChartData);
                createTimeGraph(lineChartData, connectedDotPctData);
                createSankey(sankeyData);
                createRegionTreemap(rows)
                const taxonomicFilter = document.getElementById("taxonomicFilter");
                taxonomicFilter.addEventListener("change", function () {
                    const selectedGroup = this.value;
                    createSankey(sankeyData,selectedGroup);
                });
            })
            .catch(error => {
                console.error("Error loading australia-states.json:", error);
            });
        })
        .catch(error => {
            console.error("Error loading CSV:", error);
        });

// ===============================================================================
// DATA.GOV DATASET
fetch(data_gov)
    .then(response => response.arrayBuffer())
    .then(arrayBuffer => {
        const workbook = XLSX.read(arrayBuffer, {type: "array"});
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet);
        const stateColumns = ["ACT","NSW", "NT", "QLD", "SA","TAS","VIC","WA"];
        const kingdomCounts = {};
        const normalCounts = {};
        const sunburstNodes = new Map();

        // ===============================================================================
        //SUNBURST ADD NODE FUNCTION
        function addSunburstNode(id, parent, name, level, group=null) {
            if (!sunburstNodes.has(id)) {
                sunburstNodes.set(id, {
                    id: id,
                    parent: parent,
                    name: name,
                    level: level,
                    group: group,
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
            
            if (!species || !status || !kingdom || !className ||!family ||!genus) {
                return;
            }
            // group animal and platns together
            let group;
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
                if (value.toLowerCase() === "yes") {
                    // TORNADO CHART DATA
                    const kingdomkey = `${state}|${group}`;
                    if (!kingdomCounts[kingdomkey]) {
                                kingdomCounts[kingdomkey] = {
                                    State: state,
                                    group: group,
                                    species: new Set()
                                };
                            }
                    kingdomCounts[kingdomkey].species.add(species);

                    // ===============================================================================
                    // NORMAL STACKED BAR DATA 
                    if (!normalCounts[state]) {
                        normalCounts[state] = {};
                    }
                    if (!normalCounts[state][status]) {
                        normalCounts[state][status] = new Set();
                    }
                    normalCounts[state][status].add(species);
                }
                });

            // ===============================================================================
            // SUNBURST NODE (ROOT -> KINGDOM -> CLASS -> GENUS)
            const root = addSunburstNode("root",null,"Australia","Root","All");
            const kingdomId = `kingdom|${kingdom}`;
            const kingdomNode = addSunburstNode(kingdomId,"root", kingdom,"Kingdom", group);
            const classId = `class|${kingdom}|${className}`;
            const classNode = addSunburstNode(classId, kingdomId, className, "Class", group);
            const familyId = `family|${kingdom}|${className}|${family}`;
            const familyNode = addSunburstNode(familyId, classId, family, "Family", group);
            const genusId =`genus|${kingdom}|${className}|${family}|${genus}`;
            const genusNode = addSunburstNode(genusId, familyId, genus, "Genus", group);
            // Add species to EVERY ancestor
            root.species.add(species);
            kingdomNode.species.add(species);
            classNode.species.add(species);
            familyNode.species.add(species);
            genusNode.species.add(species);
        });
        
        // ===============================================================================
        // TORNADO CHART
        const tornadochartData = Object.values(kingdomCounts).map(d => ({
            State: d.State,
            group: d.group,
            // Animals go to the left, Plants go to the right
            species_count: d.group === "Animals"? -d.species.size: d.species.size
        }));

        // ===============================================================================
        //NORMAL STACKED BAR CHART
        const normalStatusOrder = [
            "Critically Endangered",
            "Endangered",
            "Vulnerable",
            "Conservation Dependent",
            "Extinct",
            "Extinct in the wild"
        ];
        const normalStackedData = [];
        stateColumns.forEach(state => {
            const total = normalStatusOrder.reduce((sum, status) => {
                return sum + (normalCounts[state]?.[status]?.size || 0);
            }, 0);
            if (total === 0) return;
            let cumulative = 0;
            normalStatusOrder.forEach(status => {
                const count = normalCounts[state]?.[status]?.size || 0;
                if (count === 0) return;
                const percentage = (count / total) * 100;
                const startValue = cumulative;
                const endValue = cumulative + percentage;
                normalStackedData.push({
                    State: state,
                    ThreatStatus: status,
                    species_count: count,
                    percentage: percentage,
                    startValue: startValue,
                    endValue: endValue
                });
                cumulative = endValue;
            });
        });

        // ===============================================================================
        //  SUNBURST DIAGRAM
        const hierarchyNodes = Array.from(sunburstNodes.values()).map(d => ({
            id: d.id,
            parent: d.parent,
            name: d.name,
            level: d.level,
            group: d.group,
            species_count: d.species.size,
            leafValue: d.level === "Genus" ? d.species.size : 0
        }));

        // ===============================================================================
        // LOAD DATA
        fetch(MAP_FILE)
            .then(response => response.json())
            .then(mapData => {
                createChart(tornadochartData, normalStackedData, hierarchyNodes)
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
                    "species_count",
                    "standardised_count"
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

function createChart(tornadochartData, normalStackedData, hierarchyNodes) {
    fetch("Charts/tornado_chart.json")
        .then(response => response.json())
        .then(spec => {
            spec.data.values = tornadochartData;
            vegaEmbed("#tornado_chart",spec,{actions: false});
        })
        .catch(error => {console.error("tornado chart error:", error);
        });

    fetch("Charts/normal_stacked_bar.json")
        .then(response => response.json())
        .then(spec => {
            spec.data[0].values = normalStackedData;
            vegaEmbed("#normal_stacked_bar", spec, {
                actions: false,
                renderer: "svg"
            })
            .then(() => {console.log("normal stacked bar successfully displayed.");})
            .catch(error => {console.error("normal stacked bar Vega error:", error);});
        })
        .catch(error => {console.error("Error loading normal_stacked_bar.json:",error);
        });

    fetch("Charts/sunburst.json")
        .then(response => response.json())
        .then(spec => {
            spec.data[0].values = hierarchyNodes;
            vegaEmbed("#sunburst", spec, {
                actions: false,
                renderer: "svg"
            })
            .then(() => {console.log("Sunburst successfully displayed.");})
            .catch(error => {console.error("Sunburst Vega error:",error);});
        })
        .catch(error => {console.error("Error loading sunburst.json:",error);
        });
}

function createTimeGraph(lineChartData, connectedDotPctData) {
    fetch("Time-Series/multi_line_chart.json")
        .then(response => response.json())
        .then(spec => {
            spec.data.values = lineChartData;
            vegaEmbed("#multi_line",spec,{ actions: false });
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
            .then(() => {console.log("Connected dot plot successfully displayed.");})
            .catch(error => {console.error("Connected dot plot Vega error:",error);});

        })
        .catch(error => {console.error("Error loading connected_dot_plot.json:",error);
        });
}

function createBumpChart(bumpChartData) {
    fetch("Time-Series/bump_chart.json")
        .then(response => response.json())
        .then(spec => {
            if (!spec.data) {spec.data = {};}
            spec.data.values = bumpChartData;
            vegaEmbed("#bump_chart",spec,{
                    actions: false,
                    renderer: "svg"
            })
            .then(() => {console.log("Bump chart successfully displayed."); })
            .catch(error => {console.error("Bump chart Vega error:",error);});
        })
        .catch(error => {console.error( "Error loading bump_chart.json:",error);});
}

function createRegionTreemap(treemapData) {
    const container = document.getElementById("regionTreemapChart");
    const stateFilter = document.getElementById("stateFilter");
    if (!container || !stateFilter) {
        console.error("Treemap container or state filter not found.");
        return;
    }

    const states = [...new Set(
        treemapData
            .filter(d => d.depth === 1)
            .map(d => d.name)
            .filter(Boolean)
    )].sort((a, b) => a.localeCompare(b));

    stateFilter.replaceChildren();
    states.forEach(state => {
        const option = document.createElement("option");
        option.value = state;
        option.textContent = state;
        stateFilter.appendChild(option);
    });

    if (!states.length) {
        console.error("No states found in treemap data.");
        return;
    }
    stateFilter.value = states[0];

    fetch("Charts/region_treemap.json")
        .then(response => {
            if (!response.ok) {
                throw new Error("Could not load region_treemap.json");
            }
            return response.json();
        })
        .then(spec => {
            let currentView = null;

            async function renderTreemap() {
                const selectedState = stateFilter.value;
                const regions = treemapData.filter(d =>
                    d.depth === 2 &&
                    d.state === selectedState &&
                    d.region &&
                    Number(d.value) > 0
                );

                if (!regions.length) {
                    console.warn(`No region data found for ${selectedState}.`);
                    return;
                }

                const rootId = `selected:${selectedState}`;
                const regionRows = regions
                    .map((d, index) => ({
                        id: `region:${index}:${d.region}`,
                        parent: rootId,
                        name: d.region,
                        region: d.region,
                        state: selectedState,
                        value: Number(d.value)
                    }))
                    .sort((a, b) => b.value - a.value);

                const totalSpecies = regionRows.reduce(
                    (sum, d) => sum + d.value,
                    0
                );

                const filteredRows = [
                    {
                        id: rootId,
                        parent: null,
                        name: selectedState,
                        state: selectedState,
                        value: 0
                    },
                    ...regionRows
                ];

                const updatedSpec = structuredClone(spec);
                const table = updatedSpec.data.find(d => d.name === "table");
                if (!table) {
                    console.error("Vega spec has no table dataset.");
                    return;
                }

                table.values = filteredRows;
                console.table(
                    filteredRows
                        .filter(d => d.parent === rootId)
                        .map(d => ({
                            region: d.name,
                            species: d.value
                        }))
                );
                updatedSpec.title = {
                    ...updatedSpec.title,
                    text: `Endangered Species by Region — ${selectedState}`,
                    subtitle: "Rectangle area represents unique Endangered species recorded in each region"
                };
                if (currentView) {
                    currentView.finalize();
                    currentView = null;
                }
                container.replaceChildren();

                try {
                    const result = await vegaEmbed(
                        container,
                        updatedSpec,
                        { 
                            actions: false
                        }
                    );
                    currentView = result.view;
                    const treeData = result.view.data("table");

                    console.table(
                        treeData
                            .filter(d => d.parent && d.x0 !== undefined)
                            .map(d => ({
                                region: d.name,
                                species: d.value,
                                width: d.x1 - d.x0,
                                height: d.y1 - d.y0,
                                area: (d.x1 - d.x0) * (d.y1 - d.y0)
                            }))
                    );
                    await currentView.runAsync();
                } catch (error) {
                    console.error("Treemap rendering error:", error);
                }
            }

            // Window resize handler to maintain edge-to-edge scaling
            window.addEventListener("resize", () => {
                if (currentView) {
                    currentView.runAsync();
                }
            });

            stateFilter.addEventListener("change", renderTreemap);
            renderTreemap();
        })
        .catch(error => {
            console.error("Region treemap error:", error);
        });
}


// ==========================================
// CREATE SANKEY 
let rawSankeyData = [];
let sankeyView = null;
function createSankey(sankeyData) {
    rawSankeyData = sankeyData;
    console.log("Sankey raw data:", rawSankeyData);
    const filter =document.getElementById("taxonomicFilter");
    if (!filter) {
        console.warn("taxonomicFilter element was not found.");
        return;
    }
    // REGISTER LISTENER ONLY ONCE
    if (filter.dataset.sankeyListenerAttached !== "true") {
        filter.addEventListener("change", function () {
            const selectedGroup = this.value;
            console.log("Sankey selected group:", selectedGroup);
            updateSankey(selectedGroup);
        });
        filter.dataset.sankeyListenerAttached = "true";
    }
    updateSankey(filter.value);
}

// ==========================================
// UPDATE SANKEY
let sankeyRenderId = 0;
async function updateSankey(selectedGroup = "All") {
    const currentRenderId = ++sankeyRenderId;
    const filteredData =
        selectedGroup === "All"
            ? rawSankeyData
            : rawSankeyData.filter(d =>
                String(d.Taxonomic || "").trim()
                === selectedGroup
            );

    // CREATE NODES
    const nodes = [];
    const nodeMap = new Map();
    function addNode(name, level, group) {
        name = String(name || "").trim();
        if (!name) return null;
        // Keep nodes from different taxonomic groups separate
        const id = level === "group"? `group|${name}`: `${level}|${group}|${name}`;
        if (!nodeMap.has(id)) {
            const node = {
                id: id,
                name: name,
                level: level,
                group: group
            };
            nodeMap.set(id, node);
            nodes.push(node);
        }
        return id;
    }

    // CREATE LINKS
    const linkCounts = new Map();
    filteredData.forEach(row => {
        const group = String(row.Taxonomic || "").trim();
        const order = String(row.Order || "").trim();
        const family = String(row.Family || "").trim();
        if (!group ||!order ||!family) {
            return;
        }

        // CREATE NODES
        const groupId = addNode(group, "group", group);
        const orderId = addNode(order, "order", group);
        const familyId = addNode(family, "family", group);

        // GROUP → ORDER
        const groupOrderKey =`${groupId}->${orderId}`;
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
        linkCounts.get(groupOrderKey).value++;

        // ORDER → FAMILY
        const orderFamilyKey =`${orderId}->${familyId}`;
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
        linkCounts.get(orderFamilyKey).value++;
    });
    const links =Array.from(linkCounts.values());

    // HANDLE EMPTY DATA
    if (nodes.length === 0) {
        console.warn("No Sankey data for:",selectedGroup);
        return;
    }

    // COUNT NODES PER COLUMN
    const groupCount = nodes.filter(d => d.level === "group").length;
    const orderCount = nodes.filter(d => d.level === "order").length;
    const familyCount = nodes.filter(d => d.level === "family").length;
    const maxNodes = Math.max(groupCount, orderCount, familyCount);

    // DYNAMIC HEIGHT
    const topMargin = 25;
    const bottomMargin = 35;
    const dynamicHeight = Math.max(450,maxNodes * 22 + topMargin + bottomMargin);

    // CHART SIZE
    const chartWidth = 700;
    const sideMargin = 35;
    const sankeyLeft = sideMargin;
    const sankeyRight = chartWidth - sideMargin;

    // ADJUST NODE PADDING
    let nodePadding = 12;
    if (maxNodes > 10) {
        nodePadding = 9;
    }
    if (maxNodes > 20) {
        nodePadding = 6;
    }
    if (maxNodes > 30) {
        nodePadding = 4;
    }

    // SANKEY LAYOUT
    const sankeyLayout = d3.sankey()
        .nodeId(d => d.id)
        .nodeWidth(16)
        .nodePadding(nodePadding)
        .nodeAlign(d3.sankeyJustify)
        .extent([
            [sankeyLeft, topMargin],
            [sankeyRight, dynamicHeight - bottomMargin]
        ]);

    const graph =
        sankeyLayout({
            nodes:nodes.map(d => ({...d})),
            links:links.map(d => ({...d}))
        });

    const vegaNodes = graph.nodes.map(node => ({
        id: node.id,
        name: node.name,
        level: node.level,
        group: node.group,
        x0: node.x0,
        x1: node.x1,
        y0: node.y0,
        y1: node.y1,
        nodeHeight: node.y1 - node.y0
    }));

    const vegaLinks = graph.links.map(link => ({
        source: link.source.id,
        target: link.target.id,
        value: link.value,
        group: link.source.group,
        x0: link.source.x1,
        x1: link.target.x0,
        y0: link.y0,
        y1: link.y1,
        width: Math.max(1, link.width),
        path: d3.sankeyLinkHorizontal()(link)
    }));


    // LOAD VEGA JSON
    try {
        const response =await fetch("Charts/sankey_diagram.json");
        if (!response.ok) {
            throw new Error(`Failed to load sankey_diagram.json: ${response.status}`);
        }
        const spec = await response.json();
        spec.width =chartWidth;
        spec.height =dynamicHeight;

        // FIND DATASETS
        const nodeData =spec.data.find(d => d.name === "nodes");
        const linkData =spec.data.find(d => d.name === "links");
        if (!nodeData || !linkData) {
            console.error("Could not find nodes or links in sankey_diagram.json");
            return;
        }

        // INSERT DATA
        nodeData.values =vegaNodes;
        linkData.values =vegaLinks;
        spec.title.text = selectedGroup === "All"
            ? "Taxonomic Hierarchy of Threatened Species"
            : `${selectedGroup} Taxonomic Hierarchy`;

        // RENDER
        if (currentRenderId !== sankeyRenderId) {
            return;
        }
        await vegaEmbed("#sankey_diagram",spec,{actions: false, renderer: "svg"});
        // console.log("Sankey successfully displayed:", selectedGroup);
    } catch (error) {
        console.error("Sankey rendering error:", error);
    }
}