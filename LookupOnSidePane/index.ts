import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class LookupOnSidePane implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _panes: any;

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
		this.GetPanes();
		this.RenderButton();
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this._context = context;
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
		while (this._container.firstChild) {
			this._container.removeChild(this._container.firstChild);
		}
	}

	public GetPanes() {
		this._panes = (Xrm as any).App!.sidePanes;
	}

	public RenderButton() {
		let button = document.createElement("button");
		button.innerText = "→";
		button.id = "LookupOnSidePaneButton";
		button.addEventListener("click", this.OpenOnSidePane.bind(this));
		this._container.appendChild(button);
	}

	public OpenOnSidePane() {

		// IF THE COLUMN IS NULL
		if (this.HasValue()) {
			// VERIFY IF THE RECORD IS OPENED ON PANE
			if (!this.IsOpened()) {
				this._panes.createPane({
					title: this._context.parameters.lookup.raw![0].name!.toUpperCase(),
					imageSrc: this.RetrieveEntityImage(),
					alwaysRender: this._context.parameters.alwaysRender.raw == "1" ? true : false,
					canClose: this._context.parameters.canClose.raw == "1" ? true : false,
					hideHeader: this._context.parameters.hideHeader.raw == "1" ? true : false,
					paneId: this._context.parameters.lookup.raw![0].entityType + ":" + this._context.parameters.lookup.raw![0].id,
					width: this._context.parameters.width.raw!
				}).then((pane: any) => {
					pane.navigate({
						pageType: "entityrecord",
						entityName: this._context.parameters.lookup.raw![0].entityType,
						entityId: this._context.parameters.lookup.raw![0].id
					})
				});
			}
			else {
				this.Close();
			}
		}

		if (this._panes.state == 0)
			this._panes.state = 1;

		this.GetPanes();
	}

	public HasValue(): boolean {
		return this._context.parameters.lookup.raw!.length == 1
			? true
			: false;
	}

	public IsOpened(): boolean {
		return this._panes.getPane(this._context.parameters.lookup.raw![0].entityType + ":" + this._context.parameters.lookup.raw![0].id) != undefined
			? true
			: false;
	}

	public Close(): any {
		var pane = this._panes.getPane(this._context.parameters.lookup.raw![0].entityType + ":" + this._context.parameters.lookup.raw![0].id);
		if (pane !== undefined)
			pane.close();
	}

	private RetrieveEntityImage(): string | undefined {

		let icon: string | undefined;
		let req = new XMLHttpRequest();
		const baseUrl = (<any>this._context).page.getClientUrl();
		const caller = this;
		req.open("GET", baseUrl + "/api/data/v9.1/EntityDefinitions(LogicalName='" + this._context.parameters.lookup.raw![0].entityType + "')?$select=IconSmallName,ObjectTypeCode", false);
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.onreadystatechange = function () {
			if (this.readyState === 4) {
				req.onreadystatechange = null;
				if (this.status === 200) {
					var result = JSON.parse(this.response);
					if (result.ObjectTypeCode >= 10000 && result.IconSmallName != null)
						icon = baseUrl + "/WebResources/" + result.IconSmallName.toString();
					else
						icon = baseUrl + caller.GetURL(result.ObjectTypeCode);
				}
			}
		};
		req.send();

		return icon;
	}

	private GetURL(objectTypeCode: number) {

		//default icon
		var url = "/_imgs/svg_" + objectTypeCode.toString() + ".svg";

		if (!this.UrlExists(url)) {
			url = "/_imgs/ico_16_" + objectTypeCode.toString() + ".gif";

			if (!this.UrlExists(url)) {
				url = "/_imgs/ico_16_"
					+ objectTypeCode.toString() +
					".png";

				//default icon

				if (!this.UrlExists(url)) {
					url = "/_imgs/ico_16_customEntity.gif";
				}
			}
		}

		return url;
	}

	private UrlExists(url: string) {
		var http = new XMLHttpRequest();
		http.open('HEAD', url, false);
		http.send();
		return http.status != 404 && http.status != 500;
	}
}
