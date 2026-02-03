import mondaySdk from "monday-sdk-js";
import {
  mapAttachment,
  mapBoardStructure,
  mapMainBoardItem,
} from "./MondayApiMappers";
import {
  MainBoardStructure,
  MainBoardItemRaw,
  SimpleBoardItem,
  MainBoardItemPage,
  FullBoardItem,
  MainBoardItem,
  UserItem,
  DropdownOption,
} from "./types";
import { listItemsQuery, searchItemsQuery } from "./queries";

const monday = mondaySdk();
const PAGE_SIZE = 25;

let searchColumnId: string | null = null;

const MondayApi = {
  init: {
    setSearchColumn: (columnId: string) => {
      searchColumnId = columnId;
    },
  },

  getAllSimpleBoardItems: async (
    boardIds: number[]
  ): Promise<SimpleBoardItem[]> => {
    const query = `
      query {
        boards(ids: ${JSON.stringify(boardIds)}) {
          items_page(limit: 100) {
            items {
              id
              name
            }
          }
        }
      }
    `;

    const response = await monday.api(query);
    return response.data.boards.flatMap((board: any) => board.items_page.items);
  },

  getFullBoardItems: async (boardIds: number[]): Promise<FullBoardItem[]> => {
    // Csak az első board kezelése, ha több van, bővíthető
    const boardId = boardIds[0];
    const allItems: FullBoardItem[] = [];

    // Első lekérdezés
    let query = `
      query {
        boards(ids: [${boardId}]) {
          items_page(limit: 500) {
            cursor
            items {
              id
              name
              column_values {
                id
                value
                text
              }
            }
          }
        }
      }
    `;

    let response = await monday.api(query);
    let data = response.data.boards[0].items_page;

    allItems.push(...data.items);

    let cursor = data.cursor;

    // Lapozás amíg cursor van
    while (cursor) {
      query = `
        query {
          next_items_page(cursor: "${cursor}", limit: 500) {
            cursor
            items {
              id
              name
              column_values {
                id
                value
                text
              }
            }
          }
        }
      `;
      response = await monday.api(query);
      data = response.data.next_items_page;

      allItems.push(...data.items);
      cursor = data.cursor;
    }

    return allItems;
  },

  getMainBoardItems: async (
    mainBoardId: number,
    searchTerm: string,
    structure: MainBoardStructure
  ): Promise<MainBoardItemPage> => {
    const query = searchTerm?.trim() ? searchItemsQuery : listItemsQuery;

    const variables = {
      boardIds: [mainBoardId],
      pageSize: PAGE_SIZE,
      searchColumnId,
      searchTerm,
    };

    const response = await monday.api(query, { variables });
    const items: MainBoardItemRaw[] = response.data.boards[0].items_page.items;

    return {
      Cursor: response.data.boards[0].items_page.cursor,
      Items: items.map((item) => mapMainBoardItem(item, structure)),
    };
  },

  getNextPageItems: async (
    cursor: string,
    filesColumnId: string,
    structure: MainBoardStructure
  ): Promise<MainBoardItemPage> => {
    const query = `
      query {
        next_items_page(limit: ${PAGE_SIZE}, cursor: "${cursor}") {
          cursor
          items {
            id
            name
            column_values {
              id
              value
              text
            }
            assets(assets_source: columns, column_ids: ["${filesColumnId}"]) {
              id
              url
              name
              file_extension
              url_thumbnail
            }
          }
        }
      }
    `;

    const response = await monday.api(query);
    const items: MainBoardItemRaw[] = response.data.next_items_page.items;

    return {
      Cursor: response.data.next_items_page.cursor,
      Items: items.map((item) => mapMainBoardItem(item, structure)),
    };
  },

  getAdoazonositok: async (
    mainBoardId: number,
    column: string
  ): Promise<FullBoardItem[]> => {
    const query = `
      query {
        boards(ids: ${JSON.stringify(mainBoardId)}) {
          items_page {
            items {
              column_values(ids: ["${column}"]) {
                text
              }
            }
          }
        }
      }
    `;
    const response = await monday.api(query);
    return response.data.boards.flatMap((board: any) => board.items_page.items);
  },

  getMainBoardStructure: async (boardId: number): Promise<MainBoardStructure> => {
    const query = `
      query {
        boards(ids: ${boardId}) {
          columns {
            id
            type
            settings_str
          }
        }
        users {
          id
          name
          email
          photo_thumb_small
          teams {
            id
            name
          }
        }
        teams {
          id
          name
          picture_url
        }
      }
    `;

    const response = await monday.api(query);
    const columns = response.data.boards[0].columns;
    const users = response.data.users;
    const teams = response.data.teams;

    return mapBoardStructure(columns, users, teams);
  },

  getMe: async (): Promise<UserItem> => {
    const query = `
      query {
        me {
          id
          name
          email
        }
      }
    `;

    const response = await monday.api(query);
    return response.data.me;
  },

    getFilteredKoltseghelyek: async (
    boardId: number,
    columnId: string,
    searchTerm: string
    ): Promise<DropdownOption[]> => {
    if (!searchTerm || searchTerm.trim().length < 2) {
        return [];
    }

    const query = `
        query {
        boards (ids: ${boardId}) {
            items_page (query_params: {
            rules: [{
                column_id: "${columnId}",
                compare_value: ["${searchTerm}"],
                operator: contains_text
            }]
            }) {
            cursor
            items {
                id
                name
            }
            }
        }
        }
    `;

    const response = await fetch('https://monday-app-efo-export.vercel.app/api/excel_export', {
        method: 'POST',
        headers: {
        'Content-Type': 'application/json',
        },
        body: JSON.stringify({ query }),
    });

    if (!response.ok) {
        const text = await response.text();
        throw new Error(`Sikertelen lekérdezés: ${response.status} - ${text}`);
    }

    const json = await response.json();
    const items = json?.data?.boards?.[0]?.items_page?.items ?? [];

    return items.map((item: any) => ({
        caption: item.name,
        value: item.id,
        additionalInfo: "", // vagy ha van extra adat, itt adhatod meg
    })) as DropdownOption[];
    },

  changeMainBoardItem: async (
    structure: MainBoardStructure,
    mainBoardId: number,
    itemId: number,
    columnId: string,
    newValue: string
  ): Promise<void> => {
    const column = Object.values(structure).find((col) => col.id === columnId);

    let query = "";

    if (!column || (!column.isConnectedBoard && !column.isPersonOrTeam)) {
      query = `
        mutation {
          change_simple_column_value(
            item_id: ${itemId},
            board_id: ${mainBoardId},
            column_id: "${columnId}",
            value: "${newValue ?? ""}"
          ) {
            id
          }
        }
      `;
    } else {
      let value = "";
      if (column.isPersonOrTeam) {
        const kind = newValue.startsWith("T") ? "team" : "person";
        const id = newValue.substring(1);
        value = `{\\"personsAndTeams\\":[{\\"id\\":\\"${id}\\",\\"kind\\":\\"${kind}\\"}]}`;
      } else {
        value = `{\\"item_ids\\":[\\"${newValue}\\"]}`;
      }

      query = `
        mutation {
          change_multiple_column_values(
            item_id: ${itemId},
            board_id: ${mainBoardId},
            column_values: "{\\"${columnId}\\": ${value}}"
          ) {
            id
          }
        }
      `;
    }

    try {
      await monday.api(query);
    } catch (error: any) {
      console.error(error.data?.errorMessage || error.message);
    }
  },

  insMainBoardItem: async (
    structure: MainBoardStructure,
    mainBoardId: number,
    item: MainBoardItem
  ): Promise<string> => {
    const formatDateOnlyForMonday = (date: Date) =>
      date.toISOString().split("T")[0];

    const columnValueObj: Record<string, any> = {
      [import.meta.env.VITE_COLUMN_ID_SZULETESI_IDO]: {
        date: formatDateOnlyForMonday(new Date(item.Szuletesi_ido.ColumnValue)),
      },
      [import.meta.env.VITE_COLUMN_ID_ADOAZONOSITO]: item.Adoazonosito.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_EFO_IGENYLO]: item.EFO_igenylo.ColumnValue.value,
      [import.meta.env.VITE_COLUMN_ID_SZULETESI_HELY]: item.Szuletesi_hely.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_MUNKAKOR]: item.Munkakor.ColumnValue.caption,
      [import.meta.env.VITE_COLUMN_ID_SZULETESI_NEV]: item.Szuletesi_nev.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_ANYJA_NEVE]: item.Anyja_neve.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_LAKCIM]: item.Lakcim.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_ALLAMPOLGARSAG]: item.Allampolgarsag.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_TAJSZAM]: item.Tajszam.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_BANKSZAMLASZAM]: item.Bankszamlaszam.ColumnValue,
      [import.meta.env.VITE_COLUMN_ID_KOLTSEGHELY]: {
        item_ids: [item.Koltseghely.ColumnValue.value],
      },
    };

    const query = `
      mutation {
        create_item(
          board_id: ${mainBoardId},
          item_name: "${item.Name.ColumnValue}",
          column_values: "${JSON.stringify(columnValueObj).replace(/"/g, '\\"')}"
        ) {
          id
        }
      }
    `;

    try {
      const response = await fetch(
        "https://monday-app-efo-export.vercel.app/api/excel_export",
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ query }),
        }
      );

      const result = await response.json();

      if (result.errors) {
        console.error("Monday API hiba:", result.errors);
        throw new Error(JSON.stringify(result.errors));
      }

      return result.data.create_item.id;
    } catch (error: any) {
      console.error("API hívás hiba:", error.message);
      throw new Error(`Item beszúrás sikertelen: ${error.message}`);
    }
  },

  uploadFileToBoardItem: async (
    itemId: number,
    columnId: string,
    file: File
  ) => {
    const query = `
      mutation add_file($itemId: ID!, $columnId: String!, $file: File!) {
        add_file_to_column(item_id: $itemId, column_id: $columnId, file: $file) {
          id
          url
          name
          file_extension
          url_thumbnail
        }
      }
    `;

    const variables = { itemId, columnId, file };

    const res = await monday.api(query, { variables });
    return mapAttachment(res.data.add_file_to_column);
  },

  changeMainBoardItemKiszignalasCimzettje: async (
    mainBoardId: number,
    itemId: number,
    columnId: string,
    newValue: string[]
  ) => {
    const value = newValue
      .map(
        (x) =>
          `{\\"id\\":\\"${x.substring(1)}\\",\\"kind\\":\\"${
            x.startsWith("T") ? "team" : "person"
          }\\"}`
      )
      .join(",");

    const query = `
      mutation {
        change_multiple_column_values(
          item_id: ${itemId},
          board_id: ${mainBoardId},
          column_values: "{\\"${columnId}\\": {\\"personsAndTeams\\":[${value}]}}"
        ) {
          id
        }
      }
    `;

    try {
      await monday.api(query);
    } catch (error: any) {
      console.error(error.data?.errorMessage || error.message);
    }
  },
};

export default MondayApi;
