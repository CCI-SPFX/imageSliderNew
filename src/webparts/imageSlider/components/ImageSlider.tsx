








// import * as React from "react";
// import { useState, useEffect, useRef } from "react";
// import type { IImageSliderProps } from "./IImageSliderProps";
// import { Swiper, SwiperSlide } from "swiper/react";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";

// import {
//   EffectCoverflow,
//   Autoplay,
//   Navigation,
//   Pagination,
// } from "swiper/modules";

// import "swiper/swiper-bundle.css";
// // import "swiper/swiper-bundle.min.css";
// // import "swiper/css";
// // import "swiper/css/effect-coverflow";
// // import "swiper/css/navigation";
// // import "swiper/css/pagination";

// import "./style.css";

// // -----------------------------
// // TYPES
// // -----------------------------
// interface IFolderItem {
//   Name: string;
// }

// interface IFileItem {
//   Name: string;
//   ServerRelativeUrl: string;
// }

// const ImageSlider: React.FC<IImageSliderProps> = ({
//   sp,
//   selectedLibrary,
//   defaultFolder,
//   desktopSlidesCount,
//   desktopSpaceBetween,
//   desktopSliderHeight,
//   tabletSlidesCount,
//   tabletSpaceBetween,
//   tabletSliderHeight,
//   mobileSlidesCount,
//   mobileSpaceBetween,
//   mobileSliderHeight,
//   showPagination,
//   showNavigationArrows,
// }) => {
//   const [folders, setFolders] = useState<string[]>([]);
//   const [selectedFolder, setSelectedFolder] = useState<string>("");
//   const [images, setImages] = useState<string[]>([]);
//   const [libraryRoot, setLibraryRoot] = useState<string>("");

//   // Gate: true only once we've resolved which folder should load first
//   // (either the valid defaultFolder from the property pane, or "All"
//   // if no valid default exists). Nothing fetches images before this.
//   const [folderReady, setFolderReady] = useState<boolean>(false);

//   const [lightboxOpen, setLightboxOpen] = useState(false);
//   const [currentIndex, setCurrentIndex] = useState(0);

//   // Skips the defaultFolder-change effect on the very first render,
//   // since the initial folder resolution already happens inside loadFolders().
//   const isFirstRun = useRef(true);

//   // Guards against stale async responses overwriting newer ones.
//   const loadRequestId = useRef(0);

//   const openLightbox = (index: number) => {
//     setCurrentIndex(index);
//     setLightboxOpen(true);
//   };

//   const closeLightbox = () => {
//     setLightboxOpen(false);
//   };

//   const goToPrev = () => {
//     setCurrentIndex((prev) => (prev === 0 ? images.length - 1 : prev - 1));
//   };

//   const goToNext = () => {
//     setCurrentIndex((prev) => (prev === images.length - 1 ? 0 : prev + 1));
//   };

//   // -----------------------------
//   // GET LIBRARY ROOT URL
//   // -----------------------------
//   const loadLibraryRoot = async (): Promise<void> => {
//     if (!selectedLibrary) return;

//     try {
//       const lib = await sp.web.lists
//         .getByTitle(selectedLibrary)
//         .select("RootFolder/ServerRelativeUrl")
//         .expand("RootFolder")();

//       setLibraryRoot(lib.RootFolder.ServerRelativeUrl);
//     } catch (err) {
//       console.log(err);
//     }
//   };

//   // -----------------------------
//   // LOAD FOLDERS
//   // -----------------------------
//   const loadFolders = async (): Promise<void> => {
//     if (!selectedLibrary) return;

//     try {
//       const res: IFolderItem[] = await sp.web.lists
//         .getByTitle(selectedLibrary)
//         .rootFolder.folders.select("Name")();

//       const folderNames = res
//         .map((f) => f.Name)
//         .filter(
//           (name) =>
//             name !== "Forms" && !name.startsWith("_") && !name.startsWith("."),
//         );

//       setFolders(folderNames);

//       // Resolve the starting folder HERE, in the same pass as setFolders,
//       // instead of a separate effect. This avoids two competing loadImages
//       // calls (one for "All", one for defaultFolder) racing each other.
//       const initialFolder =
//         defaultFolder && folderNames.includes(defaultFolder)
//           ? defaultFolder
//           : "All";

//       setSelectedFolder(initialFolder);
//       setFolderReady(true); // now it's safe to fetch images
//     } catch (err) {
//       console.log(err);
//     }
//   };

//   // GET images from a single folder path
//   const getImages = async (path: string): Promise<string[]> => {
//     try {
//       const files: IFileItem[] = await sp.web
//         .getFolderByServerRelativePath(path)
//         .files.select("Name", "ServerRelativeUrl")();

//       return files
//         .filter((f) => /\.(jpg|jpeg|png|gif|webp)$/i.test(f.Name))
//         .map((f) => f.ServerRelativeUrl);
//     } catch {
//       return [];
//     }
//   };

//   // -----------------------------
//   // LOAD IMAGES (FILTER LOGIC)
//   // -----------------------------
//   const loadImages = async (folder: string): Promise<void> => {
//     if (!libraryRoot) return;

//     const requestId = ++loadRequestId.current;
//     let result: string[] = [];

//     if (folder === "All") {
//       // fetch root + all subfolders in parallel instead of sequentially
//       const allResults = await Promise.all([
//         getImages(libraryRoot),
//         ...folders.map((f) => getImages(`${libraryRoot}/${f}`)),
//       ]);
//       result = allResults.flat();
//     } else {
//       result = await getImages(`${libraryRoot}/${folder}`);
//     }

//     // Ignore stale/out-of-order responses â€” only the latest request wins
//     if (requestId === loadRequestId.current) {
//       setImages(result);
//     }
//   };

//   // -----------------------------
//   // EFFECTS FLOW
//   // -----------------------------

//   // Library changed -> reset ready flag and reload root
//   useEffect(() => {
//     setFolderReady(false);
//     loadLibraryRoot();
//   }, [selectedLibrary]);

//   // Root resolved -> load folders (which also resolves initial selectedFolder)
//   useEffect(() => {
//     if (libraryRoot) {
//       loadFolders();
//     }
//   }, [libraryRoot]);

//   // Only fetch images once the folder has been resolved (folderReady),
//   // so "All" is never fetched transiently before the real default folder.
//   useEffect(() => {
//     if (libraryRoot && folderReady) {
//       loadImages(selectedFolder);
//     }
//   }, [selectedFolder, folders, libraryRoot, folderReady]);

//   // React to defaultFolder changing AFTER initial load (e.g. edited in the
//   // property pane mid-session). Skipped on first run since loadFolders()
//   // already resolved the initial folder.
//   useEffect(() => {
//     if (isFirstRun.current) {
//       isFirstRun.current = false;
//       return;
//     }
//     if (folders.length === 0) return;

//     const nextFolder =
//       defaultFolder && folders.includes(defaultFolder) ? defaultFolder : "All";

//     setSelectedFolder(nextFolder);
//   }, [defaultFolder]);

//   // -----------------------------
//   // UI
//   // -----------------------------
//   return (
//     <div className="imageSlider">
//       {/* Folder Filter */}
//       <div className="filterBox">
//         <select
//           value={selectedFolder}
//           onChange={(e) => setSelectedFolder(e.target.value)}
//         >
//           <option value="All">All</option>

//           {folders.map((folder, i) => (
//             <option key={i} value={folder}>
//               {folder}
//             </option>
//           ))}
//         </select>
//       </div>

//       {/* Navigation */}
//       <div className="custom-prev">
//         <svg
//           xmlns="http://www.w3.org/2000/svg"
//           width="22"
//           height="22"
//           viewBox="0 0 24 24"
//           fill="none"
//         >
//           <path
//             d="M14.5 6L8.5 12L14.5 18"
//             stroke="currentColor"
//             stroke-width="2.5"
//             stroke-linecap="round"
//             stroke-linejoin="round"
//           />
//         </svg>
//       </div>
//       <div className="custom-next">
//         <svg
//           xmlns="http://www.w3.org/2000/svg"
//           width="22"
//           height="22"
//           viewBox="0 0 24 24"
//           fill="none"
//         >
//           <path
//             d="M9.5 6L15.5 12L9.5 18"
//             stroke="currentColor"
//             stroke-width="2.5"
//             stroke-linecap="round"
//             stroke-linejoin="round"
//           />
//         </svg>
//       </div>

//       {/* Swiper */}
//       <div
//         className="swiperWrapper"
//         style={{
//           ["--desktop-height" as any]: `${desktopSliderHeight}px`,
//           ["--tablet-height" as any]: `${tabletSliderHeight}px`,
//           ["--mobile-height" as any]: `${mobileSliderHeight}px`,
//         }}
//       >
//         <Swiper
//           key={`${desktopSlidesCount}-${desktopSpaceBetween}-${tabletSlidesCount}-${tabletSpaceBetween}-${mobileSlidesCount}-${mobileSpaceBetween}`}
//           slidesPerView={2}
//           loop={true}
//           spaceBetween={30}
//           autoplay={{ delay: 5000, disableOnInteraction: false }}
//           pagination={showPagination ? { clickable: true } : false}
//           navigation={
//             showNavigationArrows
//               ? {
//                   prevEl: ".custom-prev",
//                   nextEl: ".custom-next",
//                 }
//               : false
//           }
//           breakpoints={{
//             // Mobile (0px and up)
//             0: {
//               slidesPerView: mobileSlidesCount,
//               spaceBetween: Number(mobileSpaceBetween) || 10,
//             },
//             // Tablet (768px and up)
//             768: {
//               slidesPerView: tabletSlidesCount,
//               spaceBetween: Number(tabletSpaceBetween) || 15,
//             },
//             // Desktop (1024px and up)
//             1024: {
//               slidesPerView: desktopSlidesCount,
//               spaceBetween: Number(desktopSpaceBetween) || 20,
//             },
//           }}
//           modules={[EffectCoverflow, Autoplay, Pagination, Navigation]}
//           className="customSwiper"
//         >
//           {images.map((img, i) => (
//             <SwiperSlide key={i}>
//               <img
//                 src={img}
//                 alt={`slide-${i}`}
//                 onClick={() => openLightbox(i)}
//               />
//             </SwiperSlide>
//           ))}
//         </Swiper>
//         {/* For Popup of images  */}
//         {lightboxOpen && (
//           <div className="lightboxOverlay" onClick={closeLightbox}>
//             <button
//               className="lightboxClose"
//               onClick={(e) => {
//                 e.stopPropagation();
//                 closeLightbox();
//               }}
//               aria-label="Close"
//             >
//               &times;
//             </button>

//             <button
//               className="lightboxPrev"
//               onClick={(e) => {
//                 e.stopPropagation();
//                 goToPrev();
//               }}
//               aria-label="Previous"
//             >
//               &#10094;
//             </button>

//             <div
//               className="lightboxImageWrapper"
//               onClick={(e) => e.stopPropagation()}
//             >
//               <img
//                 src={images[currentIndex]}
//                 alt={`enlarged-${currentIndex}`}
//                 className="lightboxImage"
//               />
//             </div>

//             <button
//               className="lightboxNext"
//               onClick={(e) => {
//                 e.stopPropagation();
//                 goToNext();
//               }}
//               aria-label="Next"
//             >
//               &#10095;
//             </button>
//           </div>
//         )}
//       </div>
//     </div>
//   );
// };

// export default ImageSlider;





import * as React from "react";
import { useState, useEffect, useRef } from "react";
import type { IImageSliderProps } from "./IImageSliderProps";
import { Swiper, SwiperSlide } from "swiper/react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";

import {
  EffectCoverflow,
  Autoplay,
  Navigation,
  Pagination,
} from "swiper/modules";

import "swiper/swiper-bundle.css";
// import "swiper/swiper-bundle.min.css";
// import "swiper/css";
// import "swiper/css/effect-coverflow";
// import "swiper/css/navigation";
// import "swiper/css/pagination";

import "./style.css";

// -----------------------------
// TYPES
// -----------------------------
interface IFolderItem {
  Name: string;
}

interface IFileItem {
  Name: string;
  ServerRelativeUrl: string;
}

// -----------------------------
// MODULE-LEVEL CACHES
// -----------------------------
// These live outside the component so they survive React remounts.
// SPFx typically remounts webpart components when toggling edit/read
// mode, opening the property pane, etc. Without this cache, every
// remount re-runs the full REST call chain (library root -> folders ->
// images), which is what causes the 30-40s delay entering edit mode.
const libraryRootCache = new Map<string, string>();
const foldersCache = new Map<string, string[]>();
const imagesCache = new Map<string, string[]>();

const ImageSlider: React.FC<IImageSliderProps> = ({
  sp,
  selectedLibrary,
  defaultFolder,
  desktopSlidesCount,
  desktopSpaceBetween,
  desktopSliderHeight,
  tabletSlidesCount,
  tabletSpaceBetween,
  tabletSliderHeight,
  mobileSlidesCount,
  mobileSpaceBetween,
  mobileSliderHeight,
  showPagination,
  showNavigationArrows,
}) => {
  const [folders, setFolders] = useState<string[]>([]);
  const [selectedFolder, setSelectedFolder] = useState<string>("");
  const [images, setImages] = useState<string[]>([]);
  const [libraryRoot, setLibraryRoot] = useState<string>("");

  // Gate: true only once we've resolved which folder should load first
  // (either the valid defaultFolder from the property pane, or "All"
  // if no valid default exists). Nothing fetches images before this.
  const [folderReady, setFolderReady] = useState<boolean>(false);

  const [lightboxOpen, setLightboxOpen] = useState(false);
  const [currentIndex, setCurrentIndex] = useState(0);

  // Skips the defaultFolder-change effect on the very first render,
  // since the initial folder resolution already happens inside loadFolders().
  const isFirstRun = useRef(true);

  // Guards against stale async responses overwriting newer ones.
  const loadRequestId = useRef(0);

  const openLightbox = (index: number) => {
    setCurrentIndex(index);
    setLightboxOpen(true);
  };

  const closeLightbox = () => {
    setLightboxOpen(false);
  };

  const goToPrev = () => {
    setCurrentIndex((prev) => (prev === 0 ? images.length - 1 : prev - 1));
  };

  const goToNext = () => {
    setCurrentIndex((prev) => (prev === images.length - 1 ? 0 : prev + 1));
  };

  // -----------------------------
  // GET LIBRARY ROOT URL
  // -----------------------------
  const loadLibraryRoot = async (): Promise<void> => {
    if (!selectedLibrary) return;

    // Serve from cache instantly if we've already resolved this library
    // (e.g. this is a remount caused by entering edit mode).
    const cached = libraryRootCache.get(selectedLibrary);
    if (cached) {
      setLibraryRoot(cached);
      return;
    }

    try {
      const lib = await sp.web.lists
        .getByTitle(selectedLibrary)
        .select("RootFolder/ServerRelativeUrl")
        .expand("RootFolder")();

      const rootUrl = lib.RootFolder.ServerRelativeUrl;
      libraryRootCache.set(selectedLibrary, rootUrl);
      setLibraryRoot(rootUrl);
    } catch (err) {
      console.log(err);
    }
  };

  // -----------------------------
  // LOAD FOLDERS
  // -----------------------------
  const loadFolders = async (): Promise<void> => {
    if (!selectedLibrary) return;

    const applyInitialFolder = (folderNames: string[]): void => {
      // Resolve the starting folder HERE, in the same pass as setFolders,
      // instead of a separate effect. This avoids two competing loadImages
      // calls (one for "All", one for defaultFolder) racing each other.
      const initialFolder =
        defaultFolder && folderNames.includes(defaultFolder)
          ? defaultFolder
          : "All";

      setSelectedFolder(initialFolder);
      setFolderReady(true); // now it's safe to fetch images
    };

    // Serve from cache instantly if already fetched this session.
    const cached = foldersCache.get(selectedLibrary);
    if (cached) {
      setFolders(cached);
      applyInitialFolder(cached);
      return;
    }

    try {
      const res: IFolderItem[] = await sp.web.lists
        .getByTitle(selectedLibrary)
        .rootFolder.folders.select("Name")();

      const folderNames = res
        .map((f) => f.Name)
        .filter(
          (name) =>
            name !== "Forms" && !name.startsWith("_") && !name.startsWith("."),
        );

      foldersCache.set(selectedLibrary, folderNames);
      setFolders(folderNames);
      applyInitialFolder(folderNames);
    } catch (err) {
      console.log(err);
    }
  };

  // GET images from a single folder path
  const getImages = async (path: string): Promise<string[]> => {
    try {
      const files: IFileItem[] = await sp.web
        .getFolderByServerRelativePath(path)
        .files.select("Name", "ServerRelativeUrl")();

      return files
        .filter((f) => /\.(jpg|jpeg|png|gif|webp)$/i.test(f.Name))
        .map((f) => f.ServerRelativeUrl);
    } catch {
      return [];
    }
  };

  // -----------------------------
  // LOAD IMAGES (FILTER LOGIC)
  // -----------------------------
  const loadImages = async (folder: string): Promise<void> => {
    if (!libraryRoot) return;

    const cacheKey = `${libraryRoot}::${folder}`;
    const cached = imagesCache.get(cacheKey);
    if (cached) {
      setImages(cached);
      return;
    }

    const requestId = ++loadRequestId.current;
    let result: string[] = [];

    if (folder === "All") {
      // fetch root + all subfolders in parallel instead of sequentially
      const allResults = await Promise.all([
        getImages(libraryRoot),
        ...folders.map((f) => getImages(`${libraryRoot}/${f}`)),
      ]);
      result = allResults.reduce((acc, curr) => acc.concat(curr), []);
    } else {
      result = await getImages(`${libraryRoot}/${folder}`);
    }

    // Ignore stale/out-of-order responses - only the latest request wins
    if (requestId === loadRequestId.current) {
      imagesCache.set(cacheKey, result);
      setImages(result);
    }
  };

  // -----------------------------
  // EFFECTS FLOW
  // -----------------------------

  // Library changed -> reset ready flag and reload root
  useEffect(() => {
    setFolderReady(false);
    loadLibraryRoot();
  }, [selectedLibrary]);

  // Root resolved -> load folders (which also resolves initial selectedFolder)
  useEffect(() => {
    if (libraryRoot) {
      loadFolders();
    }
  }, [libraryRoot]);

  // Only fetch images once the folder has been resolved (folderReady),
  // so "All" is never fetched transiently before the real default folder.
  useEffect(() => {
    if (libraryRoot && folderReady) {
      loadImages(selectedFolder);
    }
  }, [selectedFolder, folders, libraryRoot, folderReady]);

  // React to defaultFolder changing AFTER initial load (e.g. edited in the
  // property pane mid-session). Skipped on first run since loadFolders()
  // already resolved the initial folder.
  useEffect(() => {
    if (isFirstRun.current) {
      isFirstRun.current = false;
      return;
    }
    if (folders.length === 0) return;

    const nextFolder =
      defaultFolder && folders.includes(defaultFolder) ? defaultFolder : "All";

    setSelectedFolder(nextFolder);
  }, [defaultFolder]);

  // -----------------------------
  // UI
  // -----------------------------
  return (
    <div className="imageSlider">
      {/* Folder Filter */}
      {/* <div className="filterBox">
        <select
          value={selectedFolder}
          onChange={(e) => setSelectedFolder(e.target.value)}
        >
          <option value="All">All</option>

          {folders.map((folder, i) => (
            <option key={i} value={folder}>
              {folder}
            </option>
          ))}
        </select>
      </div> */}

      {/* Navigation */}
      <div className="custom-prev">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="22"
          height="22"
          viewBox="0 0 24 24"
          fill="none"
        >
          <path
            d="M14.5 6L8.5 12L14.5 18"
            stroke="currentColor"
            stroke-width="2.5"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg>
      </div>
      <div className="custom-next">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="22"
          height="22"
          viewBox="0 0 24 24"
          fill="none"
        >
          <path
            d="M9.5 6L15.5 12L9.5 18"
            stroke="currentColor"
            stroke-width="2.5"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg>
      </div>

      {/* Swiper */}
      <div
        className="swiperWrapper"
        style={{
          ["--desktop-height" as any]: `${desktopSliderHeight}px`,
          ["--tablet-height" as any]: `${tabletSliderHeight}px`,
          ["--mobile-height" as any]: `${mobileSliderHeight}px`,
        }}
      >
        <Swiper
          key={`${desktopSlidesCount}-${desktopSpaceBetween}-${tabletSlidesCount}-${tabletSpaceBetween}-${mobileSlidesCount}-${mobileSpaceBetween}`}
          slidesPerView={2}
          loop={true}
          spaceBetween={30}
          autoplay={{ delay: 5000, disableOnInteraction: false }}
          pagination={showPagination ? { clickable: true } : false}
          navigation={
            showNavigationArrows
              ? {
                  prevEl: ".custom-prev",
                  nextEl: ".custom-next",
                }
              : false
          }
          breakpoints={{
            // Mobile (0px and up)
            0: {
              slidesPerView: mobileSlidesCount,
              spaceBetween: Number(mobileSpaceBetween) || 10,
            },
            // Tablet (768px and up)
            768: {
              slidesPerView: tabletSlidesCount,
              spaceBetween: Number(tabletSpaceBetween) || 15,
            },
            // Desktop (1024px and up)
            1024: {
              slidesPerView: desktopSlidesCount,
              spaceBetween: Number(desktopSpaceBetween) || 20,
            },
          }}
          modules={[EffectCoverflow, Autoplay, Pagination, Navigation]}
          className="customSwiper"
        >
          {images.map((img, i) => (
            <SwiperSlide key={i}>
              <img
                src={img}
                alt={`slide-${i}`}
                onClick={() => openLightbox(i)}
              />
            </SwiperSlide>
          ))}
        </Swiper>
        {/* For Popup of images  */}
        {lightboxOpen && (
          <div className="lightboxOverlay" onClick={closeLightbox}>
            <button
              className="lightboxClose"
              onClick={(e) => {
                e.stopPropagation();
                closeLightbox();
              }}
              aria-label="Close"
            >
              &times;
            </button>

            <button
              className="lightboxPrev"
              onClick={(e) => {
                e.stopPropagation();
                goToPrev();
              }}
              aria-label="Previous"
            >
              &#10094;
            </button>

            <div
              className="lightboxImageWrapper"
              onClick={(e) => e.stopPropagation()}
            >
              <img
                src={images[currentIndex]}
                alt={`enlarged-${currentIndex}`}
                className="lightboxImage"
              />
            </div>

            <button
              className="lightboxNext"
              onClick={(e) => {
                e.stopPropagation();
                goToNext();
              }}
              aria-label="Next"
            >
              &#10095;
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

export default ImageSlider;