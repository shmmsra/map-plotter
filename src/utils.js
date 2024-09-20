(function () {
    window.app = window.app || {};

    // Global Store
    const LocalStorage = new (class LocalStorageClass {
        storage = {};

        groupIds = [
            "global",
            "zipCodeLocationMap",
            "zipCodeDistanceMatrix",
            "zipCodeTimeMatrix",
        ];

        constructor() {
            this.groupIds.forEach((groupId) => {
                const groupData = localStorage.getItem(groupId);
                if (groupData) {
                    this.storage[groupId] = JSON.parse(groupData);
                } else {
                    this.storage[groupId] = {};
                }
            });
        }

        getItems(category, key) {
            return this.storage[category][key];
        }

        setItems(category, key, value) {
            this.storage[category][key] = value;
        }

        save() {
            this.groupIds.forEach((groupId) => {
                localStorage.setItem(
                    groupId,
                    JSON.stringify(this.storage[groupId])
                );
            });
        }

        clear() {
            this.storage = {};
            this.groupIds.forEach((groupId) => {
                localStorage.removeItem(groupId);
            });
        }
    })();

    function BatchExecutor(asyncFunc, argsList, params, batchSize) {
        return new Promise(async (resolve, reject) => {
            try {
                const results = [];
                let index = 0;

                // A helper function to execute the batch
                async function executeBatch() {
                    const currentBatch = [];

                    // Run up to `batchSize` functions in parallel
                    for (
                        let i = 0;
                        i < batchSize && index < argsList.length;
                        i++, index++
                    ) {
                        currentBatch.push(asyncFunc(argsList[index], params));
                    }

                    // Wait for all functions in the current batch to complete
                    const batchResults = await Promise.all(currentBatch);
                    results.push(...batchResults);
                }

                // Keep running batches until all arguments are processed
                while (index < argsList.length) {
                    await executeBatch();
                }

                resolve(results);
                return;
            } catch (error) {
                reject(error);
            }
        });
    }

    function ReadFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            let fileType = "";

            // When the file is successfully read
            reader.onload = function (event) {
                try {
                    if (fileType === "xlsx") {
                        const data = new Uint8Array(event.target.result);
                        const workbook = XLSX.read(data, { type: "array" });
                        // Assuming a single worksheet, hence reading from the first sheet
                        const jsonData = XLSX.utils.sheet_to_json(
                            workbook.Sheets[workbook.SheetNames[0]]
                        );
                        resolve(jsonData);
                    } else if (fileType === "json") {
                        const json = JSON.parse(event.target.result);
                        resolve(json);
                    }
                } catch (e) {
                    alert("Error parsing JSON file.");
                    reject(e);
                }
            };

            if (!file) {
                alert("Please select a Excel or JSON file first.");
                reject(Error("INVALID_FILE"));
                return;
            } else if (file.type.includes("officedocument")) {
                // Read the file as a bytes array
                reader.readAsArrayBuffer(file);
                fileType = "xlsx";
            } else if (file.type.includes("json")) {
                // Read the file as a text string
                reader.readAsText(file);
                fileType = "json";
            } else {
                alert("Please select a Excel or JSON file first.");
                reject(Error("INVALID_FILE_TYPE"));
                return;
            }
        });
    }

    function SaveFile(data, fileName, fileType) {
        return new Promise((resolve, reject) => {
            if (fileType === "json") {
                // Helper function to download file
                function downloadFile(data, filename, type) {
                    const file = new Blob([data], { type });
                    const a = document.createElement("a");
                    const url = URL.createObjectURL(file);
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    setTimeout(() => {
                        document.body.removeChild(a);
                        window.URL.revokeObjectURL(url);
                    }, 0);
                }
                const json = JSON.stringify(data, null, 4);
                downloadFile(json, fileName, "application/json");
                resolve();
            } else if (fileType === "xlsx") {
                const wb = XLSX.utils.book_new();
                Object.keys(outputData).forEach((key) => {
                    const ws = XLSX.utils.json_to_sheet(outputData[key]);
                    XLSX.utils.book_append_sheet(wb, ws, key);
                });
                XLSX.writeFile(wb, fileName);
                resolve();
            } else {
                reject(Error("INVALID_FILE_TYPE"));
                return;
            }
        });
    }

    function LoadGoogleMapsScript() {
        return new Promise((resolve, reject) => {
            /**
             * Retrieves the Google Maps API Key from LocalStorage or prompts the user.
             * @returns {Promise<string>} The API Key.
             */
            function getApiKey() {
                return new Promise((resolve, reject) => {
                    let apiKey = LocalStorage.getItems(
                        "global",
                        "GoogleMapsAPIKey"
                    );
                    if (!apiKey) {
                        apiKey = prompt(
                            "Please enter your Google Maps API Key:"
                        );
                        if (!apiKey) {
                            alert("API Key is required. Please try again.");
                            reject(Error("INVALID_API_KEY"));
                            return;
                        }
                        LocalStorage.setItems(
                            "global",
                            "GoogleMapsAPIKey",
                            apiKey
                        );
                    }
                    resolve(apiKey);
                });
            }

            getApiKey()
                .then((apiKey) => {
                    if (window.google && window.google.maps) {
                        resolve();
                        return;
                    }
                    const script = document.createElement("script");
                    script.src = `https://maps.googleapis.com/maps/api/js?key=${apiKey}&callback=initMap&libraries=geometry`;
                    script.async = true;
                    script.defer = true;
                    script.onerror = () =>
                        reject(new Error("Failed to load Google Maps script."));
                    window.initMap = () => {
                        resolve();
                    };
                    document.head.appendChild(script);
                    LocalStorage.save();
                })
                .catch(reject);
        });
    }

    function GetGeocodes(zipCode) {
        return new Promise((resolve, reject) => {
            if (!zipCode) {
                resolve(null);
                return;
            }

            // fetch(`https://geocode.maps.co/search?q=${zipCode}&api_key=66e6c223a4210765093320zbyc476a2`)
            //     .then(response => response.json())
            //     .then(data => {
            //         const latlanJSON = {
            //             zipCode: zipCode,
            //             longitude: data[0].lon,
            //             latitude: data[0].lat,
            //             address: data[0].display_name,
            //         };
            //         setTimeout(() => {
            //             resolve(latlanJSON);
            //         }, 10);
            //     })
            //     .catch(error => {
            //         console.warn("GetGeocodes: Failed to get location from zip code:", error);
            //         resolve(null);
            //     });
            // return;

            const geocoder = new google.maps.Geocoder();
            geocoder.geocode({ address: zipCode }, (results, status) => {
                try {
                    if (status === "OK" && results[0]) {
                        const location = results[0].geometry.location;
                        const latlanJSON = {
                            zipCode: zipCode,
                            longitude: location.lng(),
                            latitude: location.lat(),
                            address: results[0].formatted_address,
                        };
                        resolve(latlanJSON);
                    } else {
                        console.warn(
                            `GetGeocodes: Geocode failed for zip code ${zipCode}: ${status}`
                        );
                        resolve(null);
                    }
                } catch (error) {
                    console.warn(
                        "GetGeocodes: Failed to get location from zip code:",
                        error
                    );
                    resolve(null);
                }
            });
        });
    }

    function GetShortestDistance(sources, destinations) {
        function helperFunction(sourceDestinationPair) {
            return new Promise((resolve, reject) => {
                const { source, destination } = sourceDestinationPair;
                try {
                    const start = new google.maps.LatLng(
                        source.latitude,
                        source.longitude
                    );
                    const end = new google.maps.LatLng(
                        destination.latitude,
                        destination.longitude
                    );
                    const distance =
                        google.maps.geometry.spherical.computeDistanceBetween(
                            start,
                            end
                        );
                    resolve({
                        source,
                        destination,
                        distance,
                    });
                } catch (error) {
                    console.error(
                        "GetShortestDistance: Failed to get shortest distance:",
                        error
                    );
                    resolve(null);
                }
            });
        }

        return new Promise((resolve, reject) => {
            const promises = [];
            sources.forEach((source) => {
                destinations.forEach((destination) => {
                    promises.push(helperFunction({ source, destination }));
                });
            });
            Promise.all(promises).then((results) => {
                resolve(
                    results
                        .filter((r) => !!r)
                        .map((r) => ({
                            source: r.source.zipCode,
                            destination: r.destination.zipCode,
                            distance: r.distance,
                        }))
                );
            });
        });
    }

    function GetTravelDistanceAndTime(sources, destinations, mode = "DRIVING") {
        return new Promise((resolve, reject) => {
            const starts = sources.map(
                (loc) => new google.maps.LatLng(loc.latitude, loc.longitude)
            );
            const ends = destinations.map(
                (loc) => new google.maps.LatLng(loc.latitude, loc.longitude)
            );
            const service = new google.maps.DistanceMatrixService();
            const promises = [];
            promises.push(
                new Promise(async (resolve, reject) => {
                    try {
                        const response = await service.getDistanceMatrix({
                            origins: starts,
                            destinations: ends,
                            travelMode: mode,
                            unitSystem: google.maps.UnitSystem.METRIC,
                        });
                        resolve(response);
                    } catch (error) {
                        console.error(
                            "GetTravelDistanceAndTime: Failed to get travel details:",
                            error
                        );
                        resolve(null);
                    }
                })
            );
            Promise.all(promises).then((responses) => {
                try {
                    const results = [];
                    responses.forEach((response) => {
                        if (!response) {
                            console.error(
                                "GetTravelDistanceAndTime: Failed to get travel details."
                            );
                            return;
                        }
                        response.rows.forEach((row, i) => {
                            if (!row || !row.elements || !row.elements.length) {
                                console.error(
                                    "GetTravelDistanceAndTime: Failed to get travel details for the source."
                                );
                                return;
                            }
                            row.elements.forEach((element, j) => {
                                if (element.status === "OK") {
                                    const distance = element.distance.value;
                                    const duration = element.duration.value;
                                    results.push({
                                        source: sources[i].zipCode,
                                        destination: destinations[j].zipCode,
                                        distance,
                                        duration,
                                    });
                                } else {
                                    console.error(
                                        "GetTravelDistanceAndTime: Failed to get travel details for the destination: source:",
                                        sources[i].zipCode,
                                        ", destination:",
                                        destinations[j].zipCode
                                    );
                                }
                            });
                        });
                    });
                    resolve(results);
                } catch (error) {
                    console.error(
                        "GetTravelDistanceAndTime: Failed to get travel details:",
                        error
                    );
                    resolve(null);
                }
            });
        });
    }

    window.app.LocalStorage = LocalStorage;
    window.app.BatchExecutor = BatchExecutor;
    window.app.ReadFile = ReadFile;
    window.app.SaveFile = SaveFile;
    window.app.LoadGoogleMapsScript = LoadGoogleMapsScript;
    window.app.GetGeocodes = GetGeocodes;
    window.app.GetShortestDistance = GetShortestDistance;
    window.app.GetTravelDistanceAndTime = GetTravelDistanceAndTime;
})();
