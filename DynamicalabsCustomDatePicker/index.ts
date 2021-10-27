import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class DynamicalabsCustomDatePicker implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private context: ComponentFramework.Context<IInputs>;
	private container: HTMLDivElement;
	private firstSelectElement: HTMLSelectElement;
	private secondSelectElement: HTMLSelectElement;
	private baseDate: string;
	private dateFormat: string;

	private firstValue: string;
	private secondValue: string;
	private daySelectValues: Array<string>;
	private monthSelectValues: any;
	private yearSelectValues: Array<string>;

	private _notifyOutputChanged: () => void;

	constructor()
	{

	}

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this.context = context;
		this.container = container;
		this.baseDate = context.parameters.Field.raw && context.parameters.Field.raw.length ? context.parameters.Field.raw : "";
		this.dateFormat = context.parameters.DateFormat.raw && context.parameters.DateFormat.raw.length ? context.parameters.DateFormat.raw : "";

		switch (this.dateFormat) {
			case "0": //YYYY-MM
				this.firstValue = this.baseDate ? this.baseDate.split("-")[1] : "";
				this.secondValue = this.baseDate ? this.baseDate.split("-")[0] : "";
				break;
			case "1":  //MM/DD
				this.firstValue = this.baseDate ? this.baseDate.split("/")[0] : "";
				this.secondValue = this.baseDate ? this.baseDate.split("/")[1] : "";
				break;
		}

		this._notifyOutputChanged = notifyOutputChanged;

		this.daySelectValues = ["---", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"];
		this.monthSelectValues = { "---": "0", "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04", "May": "05", "Jun": "06", "Jul": "07", "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12" };
		this.yearSelectValues = ["---", "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018", "2017", "2016", "2015", "2014", "2013", "2012", "2011", "2010", "2009", "2008", "2007", "2006", "2005", "2004", "2003", "2002", "2001", "2000", "1999", "1998", "1997", "1996", "1995", "1994", "1993", "1992", "1991", "1990"];

		this.InitFirstDropdown();
		this.InitSecondDropdown();
	}

	private InitFirstDropdown() {
		let context = this;
		this.firstSelectElement = document.createElement("select");
		this.firstSelectElement.classList.add("dnl-first-select");

		Object.keys(this.monthSelectValues).forEach(function (el) {
			context.firstSelectElement.add(new Option(el));
		});
		let monthName = Object.keys(this.monthSelectValues).find(el => this.monthSelectValues[el] == this.firstValue) as string;
		this.firstSelectElement.selectedIndex = this.firstValue ? Object.keys(this.monthSelectValues).indexOf(monthName) : 0;

		this.firstSelectElement.addEventListener("change", this.firstChange.bind(this));
		this.container.append(this.firstSelectElement);

	}

	private InitSecondDropdown() {
		let context = this;
		this.secondSelectElement = document.createElement("select");
		this.secondSelectElement.classList.add("dnl-second-select");

		switch (this.dateFormat) {
			case "0": //YYYY-MM
				this.yearSelectValues.forEach(function (el) {
					context.secondSelectElement.add(new Option(el));
				});
				this.secondSelectElement.selectedIndex = this.secondValue ? this.yearSelectValues.indexOf(this.secondValue) : 0;
				break;
			case "1": //MM/DD
				let currentMonth = this.firstValue && this.firstValue[0] == "0" && this.firstValue.length > 1 ? this.firstValue.substring(1) : this.firstValue == "---" || !this.firstValue ? "1" : this.firstValue;
				let currentYear = new Date().getFullYear();
				let daysInMonth = this.getDaysInMonth(parseInt(currentMonth, 10), currentYear);

				for (let i = 0; i <= daysInMonth; i++) {
					context.secondSelectElement.add(new Option(this.daySelectValues[i]));
				}

				this.secondSelectElement.selectedIndex = this.secondValue ? this.daySelectValues.indexOf(this.secondValue) : 0;
				break;
		}

		this.secondSelectElement.addEventListener("change", this.secondChange.bind(this));
		this.container.append(this.secondSelectElement);
	}

	private firstChange() {
		this.firstValue = this.firstSelectElement[this.firstSelectElement.selectedIndex].innerText;
		this.firstValue = this.monthSelectValues[this.firstValue];

		switch (this.dateFormat) {
			case "1": //MM/DD
				let currentMonth = this.firstValue && this.firstValue[0] == "0" && this.firstValue.length > 1 ? this.firstValue.substring(1) : this.firstValue == "---" || !this.firstValue ? "1" : this.firstValue;
				let currentYear = new Date().getFullYear();
				let daysInMonth = this.getDaysInMonth(parseInt(currentMonth, 10), currentYear);
				let daysArray = this.daySelectValues.filter(el => { return el !== "---" && parseInt(el, 10) <= daysInMonth });
				this.clearAllOptions();

				for (let i = 0; i <= daysInMonth; i++) {
					this.secondSelectElement.add(new Option(this.daySelectValues[i]));
				}
				if (this.secondValue && this.secondValue !== "0") {
					this.secondSelectElement.selectedIndex = this.secondValue && daysArray.includes(this.secondValue) ? daysArray.indexOf(this.secondValue) + 1 : daysArray.length;
					this.secondChange();
				}

				break;
		}

		this._notifyOutputChanged();
	}

	private secondChange() {
		this.secondValue = this.secondSelectElement[this.secondSelectElement.selectedIndex].innerText.replace("---", "0");
		this._notifyOutputChanged();
	}

	private getDaysInMonth(month: number, year: number) {
		return new Date(year, month, 0).getDate();
	}

	private clearAllOptions() {
		let i = 0;
		let selectLength = this.secondSelectElement.options.length - 1;
		for (i = selectLength; i >= 0; i--) {
			this.secondSelectElement.remove(i);
		}
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.baseDate = context.parameters.Field.raw ? context.parameters.Field.raw : "";
	}

	public getOutputs(): IOutputs
	{
		let format = "";
		switch (this.dateFormat) {
			case "0": //YYYY-MM
				format = `${this.secondValue}-${this.firstValue}`;
				break;
			case "1": //MM/DD
				format = `${this.firstValue}/${this.secondValue}`
		}
		return {
			Field: (!this.secondValue || this.secondValue === "0") && (!this.firstValue || this.firstValue === "0") ? undefined : format
		};
	}

	public destroy(): void
	{}
}