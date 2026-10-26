import * as React from "react";
import * as ReactDom from "react-dom";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
// import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import ImageSlider from "./components/ImageSlider";
import { IImageSliderProps } from "./components/IImageSliderProps";

import { spfi, SPFx, SPFI } from "@pnp/sp";

// export interface IImageSliderWebPartProps {
//   selectedLibrary: string;
//   desktopSlidesCount: number;
//   desktopSpaceBetween: string;
//   desktopSliderHeight: string;
//   tabletSlidesCount: number;
//   tabletSpaceBetween: string;
//   tabletSliderHeight: string;
//   mobileSlidesCount: number;
//   mobileSpaceBetween: string;
//   mobileSliderHeight: string;
//   showPagination: boolean;
//   showNavigationArrows: boolean;
// }
// export default class ImageSliderWebPart extends BaseClientSideWebPart<IImageSliderWebPartProps> {

//   private _sp!: SPFI;
//   private _libraries: { key: string; text: string }[] = [];

//   protected async onInit(): Promise<void> {
//     await super.onInit();

//     this._sp = spfi().using(SPFx(this.context));

//     await this.loadLibraries();
//   }

//   private async loadLibraries(): Promise<void> {
//     try {
//       const lists = await this._sp.web.lists
//         .filter("BaseTemplate eq 101 and Hidden eq false")
//         .select("Title")();

//       this._libraries = lists.map(l => ({
//         key: l.Title,
//         text: l.Title
//       }));

//       // refresh property pane so dropdown updates
//       this.context.propertyPane.refresh();
//     } catch (err) {
//       console.error(err);
//     }
//   }

//   public render(): void {

//     const element: React.ReactElement<IImageSliderProps> = React.createElement(
//       ImageSlider,
//       {
//         sp: this._sp,
//         selectedLibrary: this.properties.selectedLibrary,
//         desktopSlidesCount: this.properties.desktopSlidesCount,
//         desktopSpaceBetween: this.properties.desktopSpaceBetween,
//         desktopSliderHeight: this.properties.desktopSliderHeight,
//         tabletSlidesCount: this.properties.tabletSlidesCount,
//         tabletSpaceBetween: this.properties.tabletSpaceBetween,
//         tabletSliderHeight: this.properties.tabletSliderHeight,
//         mobileSlidesCount: this.properties.mobileSlidesCount,
//         mobileSpaceBetween: this.properties.mobileSpaceBetween,
//         mobileSliderHeight: this.properties.mobileSliderHeight,
//         showPagination: this.properties.showPagination,
//         showNavigationArrows: this.properties.showNavigationArrows
//       }
//     );

//     ReactDom.render(element, this.domElement);
//   }

//   protected onDispose(): void {
//     ReactDom.unmountComponentAtNode(this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   // PROPERTY PANE
//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//   {
//     header: {
//       description: "Gallery Slider Settings"
//     },
//     groups: [
//     {
//       groupName: "Library",
//       groupFields: [
//         PropertyPaneDropdown("selectedLibrary", {
//           label: "Select Document Library",
//           options: this._libraries
//         })
//       ]
//     },
//     {
//       groupName: "Desktop Settings",
//       groupFields: [
//         PropertyPaneSlider("desktopSlidesCount", {
//           label: "Number of Slides",
//           min: 1,
//           max: 8,
//           value: 4,
//           showValue: true
//         }),
//         PropertyPaneTextField("desktopSpaceBetween", {
//           label: "Space Between Slides (px)",
//           value: "20"
//         }),
//         PropertyPaneTextField("desktopSliderHeight", {
//           label: "Slider Height (px)",
//           value: "400"
//         })
//       ]
//     },
//     {
//       groupName: "Tablet Settings",
//       groupFields: [
//         PropertyPaneSlider("tabletSlidesCount", {
//           label: "Number of Slides",
//           min: 1,
//           max: 6,
//           value: 2,
//           showValue: true
//         }),
//         PropertyPaneTextField("tabletSpaceBetween", {
//           label: "Space Between Slides (px)",
//           value: "15"
//         }),
//         PropertyPaneTextField("tabletSliderHeight", {
//           label: "Slider Height (px)",
//           value: "350"
//         })
//       ]
//     },
//     {
//       groupName: "Mobile Settings",
//       groupFields: [
//         PropertyPaneSlider("mobileSlidesCount", {
//           label: "Number of Slides on Mobile",
//           min: 1,
//           max: 3,
//           value: 1,
//           showValue: true
//         }),
//         PropertyPaneTextField("mobileSpaceBetween", {
//           label: "Space Between Slides (px)",
//           value: "10"
//         }),
//         PropertyPaneTextField("mobileSliderHeight", {
//           label: "Slider Height (px)",
//           value: "250"
//         })
//       ]
//     },
//     {
//   groupName: "Navigation & Pagination",
//   groupFields: [
//     PropertyPaneToggle("showPagination", {
//       label: "Show Pagination Dots",
//       onText: "Show",
//       offText: "Hide",
//       checked: true
//     }),
//     PropertyPaneToggle("showNavigationArrows", {
//       label: "Show Navigation Arrows",
//       onText: "Show",
//       offText: "Hide",
//       checked: true
//     })
//   ]
// }

//     ]
//   }
// ]
//     };
//   }
// }

export interface IImageSliderWebPartProps {
  selectedLibrary: string;
  defaultFolder: string; // NEW
  desktopSlidesCount: number;
  desktopSpaceBetween: string;
  desktopSliderHeight: string;
  tabletSlidesCount: number;
  tabletSpaceBetween: string;
  tabletSliderHeight: string;
  mobileSlidesCount: number;
  mobileSpaceBetween: string;
  mobileSliderHeight: string;
  showPagination: boolean;
  showNavigationArrows: boolean;
}

export default class ImageSliderWebPart extends BaseClientSideWebPart<IImageSliderWebPartProps> {
  private _sp!: SPFI;
  private _libraries: { key: string; text: string }[] = [];
  private _folders: { key: string; text: string }[] = []; // NEW

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
    await this.loadLibraries();

    // if a library was already selected from a previous save, load its folders too
    if (this.properties.selectedLibrary) {
      await this.loadFolders(this.properties.selectedLibrary);
    }
  }

  private async loadLibraries(): Promise<void> {
    try {
      const lists = await this._sp.web.lists
        .filter("BaseTemplate eq 101 and Hidden eq false")
        .select("Title")();

      this._libraries = lists.map((l) => ({ key: l.Title, text: l.Title }));
      this.context.propertyPane.refresh();
    } catch (err) {
      console.error(err);
    }
  }

  // NEW: fetch folders for the selected library
  private async loadFolders(libraryTitle: string): Promise<void> {
    if (!libraryTitle) {
      this._folders = [];
      return;
    }

    try {
      const res = await this._sp.web.lists
        .getByTitle(libraryTitle)
        .rootFolder.folders.select("Name")();

      this._folders = res
        .map((f: { Name: string }) => f.Name)
        .filter(
          (name: string) =>
            name !== "Forms" && !name.startsWith("_") && !name.startsWith("."),
        )
        .map((name: string) => ({ key: name, text: name }));

      this.context.propertyPane.refresh();
    } catch (err) {
      console.error(err);
      this._folders = [];
    }
  }

  // NEW: detect when selectedLibrary changes in the property pane
  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any,
  ): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === "selectedLibrary" && oldValue !== newValue) {
      // reset default folder since library changed
      this.properties.defaultFolder = "";
      await this.loadFolders(newValue);
      this.render();
    }
  }

  public render(): void {
    const element: React.ReactElement<IImageSliderProps> = React.createElement(
      ImageSlider,
      {
        sp: this._sp,
        selectedLibrary: this.properties.selectedLibrary,
        defaultFolder: this.properties.defaultFolder, // NEW
        desktopSlidesCount: this.properties.desktopSlidesCount,
        desktopSpaceBetween: this.properties.desktopSpaceBetween,
        desktopSliderHeight: this.properties.desktopSliderHeight,
        tabletSlidesCount: this.properties.tabletSlidesCount,
        tabletSpaceBetween: this.properties.tabletSpaceBetween,
        tabletSliderHeight: this.properties.tabletSliderHeight,
        mobileSlidesCount: this.properties.mobileSlidesCount,
        mobileSpaceBetween: this.properties.mobileSpaceBetween,
        mobileSliderHeight: this.properties.mobileSliderHeight,
        showPagination: this.properties.showPagination,
        showNavigationArrows: this.properties.showNavigationArrows,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Gallery Slider Settings" },
          groups: [
            {
              groupName: "Library",
              groupFields: [
                PropertyPaneDropdown("selectedLibrary", {
                  label: "Select Document Library",
                  options: this._libraries,
                }),
                // NEW dropdown — only shows folders once a library is selected
                PropertyPaneDropdown("defaultFolder", {
                  label: "Select Default Folder",
                  options: [
                    { key: "", text: "All (no default folder)" },
                    ...this._folders,
                  ],
                  disabled: !this.properties.selectedLibrary,
                }),
              ],
            },
            {
              groupName: "Desktop Settings",
              groupFields: [
                PropertyPaneSlider("desktopSlidesCount", {
                  label: "Number of Slides",
                  min: 1,
                  max: 8,
                  value: 4,
                  showValue: true,
                }),
                PropertyPaneTextField("desktopSpaceBetween", {
                  label: "Space Between Slides (px)",
                  value: "20",
                }),
                PropertyPaneTextField("desktopSliderHeight", {
                  label: "Slider Height (px)",
                  value: "400",
                }),
              ],
            },
            {
              groupName: "Tablet Settings",
              groupFields: [
                PropertyPaneSlider("tabletSlidesCount", {
                  label: "Number of Slides",
                  min: 1,
                  max: 6,
                  value: 2,
                  showValue: true,
                }),
                PropertyPaneTextField("tabletSpaceBetween", {
                  label: "Space Between Slides (px)",
                  value: "15",
                }),
                PropertyPaneTextField("tabletSliderHeight", {
                  label: "Slider Height (px)",
                  value: "350",
                }),
              ],
            },
            {
              groupName: "Mobile Settings",
              groupFields: [
                PropertyPaneSlider("mobileSlidesCount", {
                  label: "Number of Slides on Mobile",
                  min: 1,
                  max: 3,
                  value: 1,
                  showValue: true,
                }),
                PropertyPaneTextField("mobileSpaceBetween", {
                  label: "Space Between Slides (px)",
                  value: "10",
                }),
                PropertyPaneTextField("mobileSliderHeight", {
                  label: "Slider Height (px)",
                  value: "250",
                }),
              ],
            },
            {
              groupName: "Navigation & Pagination",
              groupFields: [
                PropertyPaneToggle("showPagination", {
                  label: "Show Pagination Dots",
                  onText: "Show",
                  offText: "Hide",
                  checked: true,
                }),
                PropertyPaneToggle("showNavigationArrows", {
                  label: "Show Navigation Arrows",
                  onText: "Show",
                  offText: "Hide",
                  checked: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
