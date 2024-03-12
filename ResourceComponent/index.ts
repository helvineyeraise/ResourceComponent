/// <reference lib="dom" />
import { resourceUsage } from "process";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
export class ResourceComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _context: ComponentFramework.Context<IInputs>;
    private _recordId: string | null = null;
    private _container: HTMLDivElement;
    private startDate: Date
    private endDate: Date
    private searchIconUrl: string;
    /**
     * Empty constructor.
     */
    constructor() {

    }
    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
        this._recordId = context.parameters.resourceProperty.raw
        if (this._context.parameters.start.raw != null) {
            this.startDate = this._context.parameters.start.raw;
        }
        if (this._context.parameters.end.raw != null) {
            this.endDate = this._context.parameters.end.raw;
        }
        //search container
        const searchContainer = document.createElement("div");
        searchContainer.id = "searchContainer";
        const searchInput = document.createElement("input");
        searchInput.type = "text";
        searchInput.id = "searchInput";
        searchInput.placeholder = "Search...";
        const searchButton = document.createElement("button");
        searchButton.id = "searchButton";
        const searchIcon = document.createElement("img");
        searchIcon.src = this.searchIconUrl;
        searchIcon.alt = "Search";
        searchButton.appendChild(searchIcon);
        searchContainer.appendChild(searchInput);
        searchContainer.appendChild(searchButton);
        this._container.appendChild(searchContainer);
        searchButton.addEventListener("click", performSearch);

        // Function to handle search
        function performSearch() {
            const searchTerm = searchInput.value;
            console.log("Performing search for:", searchTerm);
        }

        // Create button element
        const button = document.createElement("button");
        button.id = "submitBtn";
        button.textContent = "Submit";
        button.onclick = this.ReadHTMLTable;
        this._container.appendChild(button);

        // Create table element
        const table = document.createElement("table");
        table.id = "myTable";
        const thead = document.createElement("thead");
        const firstRow = document.createElement("tr");
        const firstDateHeader = document.createElement("th");
        firstDateHeader.textContent = "First Date";
        firstRow.appendChild(firstDateHeader);
        const startHeader = document.createElement("th");
        startHeader.colSpan = 35;
        startHeader.id = "start";
        firstRow.appendChild(startHeader);
        const endDateRow = document.createElement("tr");
        const lastDateHeader = document.createElement("th");
        lastDateHeader.textContent = "Last Date";
        endDateRow.appendChild(lastDateHeader);
        const endHeader = document.createElement("th");
        endHeader.colSpan = 35;
        endHeader.id = "end";
        endDateRow.appendChild(endHeader);
        const yearRow = document.createElement("tr");
        thead.appendChild(firstRow);
        thead.appendChild(endDateRow);
        thead.appendChild(yearRow);
        const tbody = document.createElement("tbody");
        tbody.id = "tbody";
        table.appendChild(thead);
        table.appendChild(tbody);
        this._container.appendChild(table);
        //Create month row based on date
        this.MonthRow();
        //Get data from Table
        this.retrieveData();
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
        this.MonthRow();
    }

    /**
     * name
     */
    public async retrieveData() {
        let arraydata=[];
        if (this._context.parameters.projectId.raw != null) {
            try {
                const result = await this._context.webAPI.retrieveMultipleRecords("eye_resourcemanagement", `?$filter=_eye_resource_value ne null and _eye_project_value eq '${this._context.parameters.projectId.raw}'`);
                if (result.entities.length > 0) {
                    arraydata.push(this.groupBy(result.entities, "_eye_resource_value"));
                    console.log(arraydata);
                    for (let i = 0; i < arraydata.length; i++) {
                        this.TableRow(arraydata[i]);
                    }
                }
            } catch (error) {
                console.error("Error retrieving data:", error);
                // Handle error
            }
        }
    }


    public TableRow(row: any[]) {
        console.log(row);
        const startYear: number = this.startDate.getFullYear();
        const endYear: number = this.endDate.getFullYear();
        const startMonth: number = this.startDate.getMonth();
        const endMonth: number = this.endDate.getMonth();
        let monthcell: string = "";
        try {
            for (let year = startYear; year <= endYear; year++) {
                const dataArray = row.find((data) => data.eye_year == year);
                console.log(dataArray);
                monthcell += `<td id="uuid" style="display:none">${dataArray.eye_resourcemanagementid}</td>`;

                for (let month = (year === startYear ? startMonth : 0); month <= (year === endYear ? endMonth : 11); month++) {
                    const monthName: string = new Date(year, month, 1).toLocaleString("default", { month: "short" });
                    if (monthName == "Jan" && year == dataArray.eye_year) {
                        monthcell += `<td contenteditable="true" id="${monthName}">${dataArray.eye_january != null ? dataArray.eye_january : 0}</td>`;
                    }
                    // Other months follow the same pattern
                }
                monthcell += `<td id="endRow" style="display:none"></td>`;
            }

            if (monthcell != null && monthcell != "") {
                const resourcecell: string = `<td>${row[0]["_eye_resource_value@OData.Community.Display.V1.FormattedValue"]}</td>`;
                const monthrow: string = `<tr>${resourcecell + monthcell}</tr>`;
                document.getElementById("tbody")!.innerHTML += monthrow;
            }
        } catch (error) {
            console.error("Error on TableRow function:", error);
        }
    }

    public MonthRow() {
        document.getElementById("start")!.innerHTML = "";
        document.getElementById("end")!.innerHTML = "";
        document.getElementById("tbody")!.innerHTML = "";
        const formattedStartDate: string = this.startDate.toLocaleDateString("en-US", {
            weekday: "short",
            month: "short",
            day: "numeric",
            year: "numeric",
        });

        const formattedEndDate: string = this.endDate.toLocaleDateString("en-US", {
            weekday: "short",
            month: "short",
            day: "numeric",
            year: "numeric",
        });
        document.getElementById("start")!.innerHTML = formattedStartDate;
        document.getElementById("end")!.innerHTML = formattedEndDate;
        const startYear: number = this.startDate.getFullYear();
        const endYear: number = this.endDate.getFullYear();
        const startMonth: number = this.startDate.getMonth();
        const endMonth: number = this.endDate.getMonth();
        let monthcell: string = "";

        for (let year = startYear; year <= endYear; year++) {
            for (let month = (year === startYear ? startMonth : 0); month <= (year === endYear ? endMonth : 11); month++) {
                const monthName: string = new Date(year, month, 1).toLocaleString("default", { month: "short" });
                monthcell += `<td>${monthName} ${year}</td>`;
            }
        }

        if (monthcell != null && monthcell != "") {
            const resourceCell: string = "<td>Resource Name</td>";
            const monthRow: string = `<tr>${resourceCell + monthcell}</tr>`;
            document.getElementById("tbody")!.innerHTML += monthRow;
        }
    }
    // public groupBy(array: any[], key: string) {
    //     return Object.values(array.reduce((result: any, currentValue: any) => {
    //         (result[currentValue[key]] = result[currentValue[key]] || []).push(currentValue);
    //         return result;
    //     }, {}));
    // }
    public groupBy(data: any[], key: string) {
        const groups: any[] = [];
        data.forEach((item: any) => {
            const value = item[key];
            const existingGroup = groups.find(group => group[key] === value);
            if (existingGroup) {
                existingGroup.items.push(item);
            } else {
                groups.push({ [key]: value, items: [item] });
            }
        });
    
        return groups;
    }

    public ReadHTMLTable() {
        let resourceid: string = "";
        let jan: number = 0;
        let feb: number = 0;
        let mar: number = 0;
        let apr: number = 0;
        let may: number = 0;
        let jun: number = 0;
        let jul: number = 0;
        let aug: number = 0;
        let sep: number = 0;
        let oct: number = 0;
        let nov: number = 0;
        let dec: number = 0;
        const oTable: HTMLTableElement | null = document.getElementById('myTable') as HTMLTableElement;
        if (oTable) {
            const rowLength: number = oTable.rows.length;
            if (rowLength > 3) {
                for (let i = 0; i < rowLength; i++) {
                    //const oCells: HTMLCollectionOf<HTMLTableDataCellElement> = oTable.rows.item(i).cells;
                    const rowItem = oTable.rows.item(i);
                    // eslint-disable-next-line no-undef
                    const oCells: HTMLCollectionOf<HTMLTableCellElement> | null = rowItem ? rowItem.cells : null;
                    let cellLength: number = 0;
                    if (oCells) {
                        cellLength = oCells.length;
                        for (let j = 0; j < cellLength; j++) {
                            if (oCells[j].id == "endRow") {
                                const resData = {
                                    "eye_january": parseFloat(jan.toString()),
                                    "eye_february": parseFloat(feb.toString()),
                                    "eye_march": parseFloat(mar.toString()),
                                    "eye_april": parseFloat(apr.toString()),
                                    "eye_may": parseFloat(may.toString()),
                                    "eye_june": parseFloat(jun.toString()),
                                    "eye_july": parseFloat(jul.toString()),
                                    "eye_august": parseFloat(aug.toString()),
                                    "eye_september": parseFloat(sep.toString()),
                                    "eye_october": parseFloat(oct.toString()),
                                    "eye_november": parseFloat(nov.toString()),
                                    "eye_december": parseFloat(dec.toString())
                                };
                                this.UpdateResource(resourceid, resData);
                                jan = 0;
                                feb = 0;
                                mar = 0;
                                apr = 0;
                                may = 0;
                                jun = 0;
                                jul = 0;
                                aug = 0;
                                sep = 0;
                                oct = 0;
                                nov = 0;
                                dec = 0;
                                resourceid = "";
                            }
                            if (oCells[j].id == "uuid") {
                                resourceid = oCells.item(j)!.innerHTML;
                            }
                            if (oCells[j].id == "Jan") {
                                jan = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Feb") {
                                feb = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Mar") {
                                mar = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Apr") {
                                apr = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "May") {
                                may = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Jun") {
                                jun = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Jul") {
                                jul = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Aug") {
                                aug = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Sep") {
                                sep = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Oct") {
                                oct = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Nov") {
                                nov = parseFloat(oCells.item(j)!.innerHTML);
                            }
                            if (oCells[j].id == "Dec") {
                                dec = parseFloat(oCells.item(j)!.innerHTML);
                            }
                        }
                    }
                }
            }
        }
    }
    public UpdateResource(resid: string, resData: any) {
        const cdsWebApi = this._context.webAPI;
        if (resid != null && resid != "") {
            cdsWebApi.updateRecord("eye_resourcemanagement", resid, resData).then(
                (response) => {
                    console.log("Record created successfully. Record ID: " + response.id);
                },
                (error) => {
                    console.error("Error creating record: " + error.message);
                }
            );
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

}
