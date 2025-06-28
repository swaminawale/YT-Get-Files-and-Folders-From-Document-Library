import * as React from "react";
import styles from "./GetFilesAndFolders.module.scss";
import type { IGetFilesAndFoldersProps } from "./IGetFilesAndFoldersProps";
import { PrimaryButton } from "@fluentui/react";
import {
  getDocumentLibraryFolderFromAbsolutePath,
  getDocumentLibraryFolderFromServerRelativePath,
  getFilesAndFoldersFromDocumentLibrary,
} from "../pnp-services/PnPService";

interface IGetFilesAndFoldersState {
  webFolders: any[];
  listFolders: any[];
  itemFolders: any[];
  loading: boolean;
  error?: string;
}

// interface ISPFolder {
//   "odata.type": string;
//   "odata.id": string;
//   "odata.editLink": string;
//   Exists: boolean;
//   ExistsAllowThrowForPolicyFailures: boolean;
//   ExistsWithException: boolean;
//   IsWOPIEnabled: boolean;
//   ItemCount: number;
//   Name: string;
//   ProgID: string | null;
//   ServerRelativeUrl: string;
//   TimeCreated: string; // ISO date string
//   TimeLastModified: string; // ISO date string
//   UniqueId: string; // GUID
//   WelcomePage: string;
// }

export default class GetFilesAndFolders extends React.Component<
  IGetFilesAndFoldersProps,
  IGetFilesAndFoldersState
> {
  constructor(props: IGetFilesAndFoldersProps) {
    super(props);
    this.state = {
      webFolders: [],
      listFolders: [],
      itemFolders: [],
      loading: false,
    };
  }

  /** One stable instance, bound with class-field syntax */
  private getFilesAndFolders = async (): Promise<void> => {
    this.setState({ loading: true, error: undefined });

    try {
      const { webFolders, listFolders, itemFolders } =
        await getFilesAndFoldersFromDocumentLibrary(
          "Documents",
          this.props.context,
          1
        );

      this.setState({ webFolders, listFolders, itemFolders });
    } catch (err: any) {
      console.error(err);
      this.setState({ error: err.message ?? "Unknown error" });
    } finally {
      this.setState({ loading: false });
    }
  };

  private getFolderFromServerRelativePath = async (): Promise<void> => {
    this.setState({ loading: true, error: undefined });

    try {
      const folder = await getDocumentLibraryFolderFromServerRelativePath(
        "/sites/<site-name>/Shared Documents",
        this.props.context
      );

      console.log(folder);
    } catch (err: any) {
      console.error(err);
      this.setState({ error: err.message ?? "Unknown error" });
    } finally {
      this.setState({ loading: false });
    }
  };

  private getFolderFromAbsolutePath = async (): Promise<void> => {
    this.setState({ loading: true, error: undefined });

    try {
      const folder = await getDocumentLibraryFolderFromAbsolutePath(
        "https://<your-tenant>.sharepoint.com/sites/<site-name>/Shared%20Documents",
        this.props.context
      );

      console.log(folder);
    } catch (err: any) {
      console.error(err);
      this.setState({ error: err.message ?? "Unknown error" });
    } finally {
      this.setState({ loading: false });
    }
  };

  public render(): React.ReactElement<IGetFilesAndFoldersProps> {
    const { webFolders, listFolders, itemFolders, loading, error } = this.state;

    return (
      <section className={styles.getFilesAndFolders}>
        <div>
          <h1>Get Files and Folders from document library</h1>

          <PrimaryButton
            text={loading ? "Loading…" : "Get SP.Web Files and Folders"}
            onClick={this.getFilesAndFolders}
            disabled={loading}
          />

          {error && <p style={{ color: "red" }}>{error}</p>}

          <div>
            <h2>Web Folders</h2>

            <pre>
              {webFolders.length
                ? JSON.stringify(webFolders, null, 2)
                : "No web folders found."}
            </pre>
          </div>

          <div>
            <h2>List Folders</h2>
            <pre>
              {listFolders.length
                ? JSON.stringify(listFolders, null, 2)
                : "No list folders found."}
            </pre>
          </div>

          <div>
            <h2>Item Folders</h2>
            <pre>
              {itemFolders.length
                ? JSON.stringify(itemFolders, null, 2)
                : "No item folders found."}
            </pre>
          </div>
        </div>

        <div>
          <h1>Get Files and Folders from server relative path</h1>
          <PrimaryButton
            text={"Get Files and Folders from Server Relative Path"}
            onClick={this.getFolderFromServerRelativePath}
            disabled={loading}
          />
        </div>

        <div>
          <h1>Get Files and Folders from absolute path</h1>
          <PrimaryButton
            text={"Get Files and Folders from Absolute Path"}
            onClick={this.getFolderFromAbsolutePath}
            disabled={loading}
          />
        </div>
      </section>
    );
  }
}
