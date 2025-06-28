import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  folderFromAbsolutePath,
  folderFromServerRelativePath,
  IFolder,
} from "@pnp/sp/folders";

export const getFilesAndFoldersFromDocumentLibrary = async (
  libraryName: string,
  context: WebPartContext,
  specificFolderId: number
) => {
  const sp = spfi().using(SPFx(context));

  //Site-level folders (e.g. "Site Assets", "Style Library", …)
  const webFolders = await sp.web.folders();

  // Library root-level sub-folders
  const listFolders = await sp.web.lists
    .getByTitle(libraryName)
    .rootFolder.folders();

  // Specific folder from the library
  const itemFolders = await sp.web.lists
    .getByTitle(libraryName)
    .items.getById(specificFolderId)
    .folder.folders();

  return { webFolders, listFolders, itemFolders };
};

export const getDocumentLibraryFolderFromServerRelativePath = async (
  serverRelativePath: string,
  context: WebPartContext
) => {
  const sp = spfi().using(SPFx(context));
  const folder: IFolder = folderFromServerRelativePath(
    sp.web,
    serverRelativePath
  );

  const folderInfo = await folder(); // metadata only
  const files = await folder.files(); // this works

  console.log("Folder Info:", folderInfo);
  console.log("Files:", files);
};

export const getDocumentLibraryFolderFromAbsolutePath = async (
  absolutePath: string,
  context: WebPartContext
) => {
  const sp = spfi().using(SPFx(context));
  const folder: IFolder = await folderFromAbsolutePath(sp.web, absolutePath);

  const folderInfo = await folder(); // metadata only
  const files = await folder.files(); // this works

  console.log("Folder Info:", folderInfo);
  console.log("Files:", files);
};
