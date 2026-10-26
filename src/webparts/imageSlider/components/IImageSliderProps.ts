import { SPFI } from "@pnp/sp";

export interface IImageSliderProps {
  // imageFolder: string; 
  sp:SPFI;
  selectedLibrary:string;
  defaultFolder: string;  
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
