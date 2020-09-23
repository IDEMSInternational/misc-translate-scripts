import * as xlsx from "xlsx";
import * as fs from "fs";
import * as path from "path";
import { exit } from "process";
import { RapidProFlowExport } from "./rapid-pro-export.model";
import { v4 as uuidv4 } from 'uuid';

const inputFolderPath = "./input";
const spreadsheetFile = "plh_input1.xlsx";

interface ConversationExcelRow {
    Type: 'Start_new_flow' | 'Send_message',
    From?: number | string,
    Condition?: string,
    Condition_Var?: string,
    MessageText: string,
    Media?: string,
    Choice_1?: string,
    Choice_2?: string,
    Choice_3?: string,
    Choice_4?: string,
    Choice_5?: string,
    Choice_6?: string,
    Choice_7?: string,
    Choice_8?: string,
    Choice_9?: string,
    Choice_10?: string,
    NodeUUID?: string,
    rapidProNode?: RapidProFlowExport.Node
}

let workbook = xlsx.readFile(path.join(inputFolderPath, spreadsheetFile));

console.log("Sheet names", workbook.SheetNames);

if (!workbook.Sheets["Content_List"]) {
    console.error("No content list!");
    exit();
}

let contentList: {
    Sheet_Name: string,
    Content_Type: "Conversation",
    Character?: "Friend" | "Guide",
    Entry_Condition?: string
}[] = xlsx.utils.sheet_to_json(workbook.Sheets["Content_List"]);

console.log("Content list parsed: ", contentList);

function getFromRowNumbers(row: ConversationExcelRow, rows: ConversationExcelRow[]): number[] {
    if (row.From) {
        if (typeof row.From === "number") {
            return [row.From];
        }
        if (typeof row.From === "string") {
            return row.From.split(",")
                .map((fromStr) => Number.parseInt(fromStr))
                .filter((num) => num != NaN)
        }
    }
    return [rows.indexOf(row) + 2];
}

function getFromNodes(row: ConversationExcelRow, rows: ConversationExcelRow[]) {
    return getFromRowNumbers(row, rows)
        .map((rowNumber) => rows[rowNumber - 2].rapidProNode)
        .filter((node) => node !== undefined);
}

function getRouterFromRow(row: ConversationExcelRow, rows: ConversationExcelRow[], nodesById: { [nodeId: string]: RapidProFlowExport.Node }) {
    const fromRowsThatExitToRouterNodes = getFromNodes(row, rows)
        .filter((node) => node !== undefined && node.exits.length > 0 && node.exits[0].destination_uuid)
        .map((node) => nodesById[node.exits[0].destination_uuid])
        .filter((node) => node.router !== undefined);
        
    return fromRowsThatExitToRouterNodes[0];
}

function processContactVariableRow(
    row: ConversationExcelRow,
    rows: ConversationExcelRow[],
    nodesById: { [nodeId: string]: RapidProFlowExport.Node },
    flow: RapidProFlowExport.Flow
) {
    let existingRouterNode = getRouterFromRow(row, rows, nodesById);
    if (!existingRouterNode) {
        let nodeId = uuidv4();
        let exits: RapidProFlowExport.Exit[] = [{
            uuid: uuidv4(),
            destination_uuid: null
        }];
        let categories: RapidProFlowExport.Category[] = [{
            exit_uuid: exits[0].uuid,
            name: row.Condition,
            uuid: uuidv4()
        }];
        let cases: RapidProFlowExport.Case[] = [{
            arguments: [row.Condition], // Multi condition handling later
            category_uuid: categories[0].uuid,
            type: "has_any_word", // Add more case handling later
            uuid: uuidv4()
        }];

        let routerNode: RapidProFlowExport.Node = {
            "uuid": nodeId,
            "actions": [],
            "router": {
                "type": "switch",
                "default_category_uuid": null, // How do we want excel to specify default?
                "cases": cases,
                "categories": categories,
                "operand": "@fields." + row.Condition_Var
            },
            "exits": exits
        };
        flow.nodes.push(routerNode);
    } else {
        existingRouterNode.router.cases.push();
    }
}

for (let contentListItem of contentList) {
    if (!workbook.Sheets[contentListItem.Sheet_Name]) {
        console.warn("No worksheet for=" + contentListItem.Sheet_Name + ":");
        break;
    }
    if (contentListItem.Content_Type === "Conversation") {
        const rows: ConversationExcelRow[] = xlsx.utils.sheet_to_json(workbook.Sheets[contentListItem.Sheet_Name]);
        console.log("Found Conversation sheet for ", contentListItem.Sheet_Name);
        let flow: RapidProFlowExport.Flow = {
            name: contentListItem.Sheet_Name,
            uuid: contentListItem.Sheet_Name + "_id",
            spec_version: "v1",
            language: "english",
            type: "type",
            nodes: [],
            _ui: null,
            revision: 0,
            expire_after_minutes: 60,
            metadata: {
                revision: 0,
            },
            localization: null
        };
        let nodesById: { [nodeId: string]: RapidProFlowExport.Node } = {};
        for (let rowNumber = 2; rowNumber < rows.length + 2; rowNumber++) {
            let row = rows[rowNumber - 2];
            console.log("row ", rowNumber, row);
            if (row.Condition) {
                if (row.Condition_Var) {
                    processContactVariableRow(row, rows, nodesById, flow);
                } else {
                    // Switch on last message choice
                }
            }
            if (row.Type === "Send_message") {
                let quick_replies = [];
                for (var i = 1; i <= 10; i++) {
                    if (row["Choice_" + i]) {
                        quick_replies.push(row["Choice_" + 1]);
                    }
                }
                let nodeId = uuidv4();

                let exits: RapidProFlowExport.Exit[] = [{
                    uuid: uuidv4(),
                    destination_uuid: undefined
                }];
                let sendMessageNode: RapidProFlowExport.Node = {
                    "uuid": nodeId,
                    "actions": [
                        {
                            "attachments": [],
                            "text": row.MessageText,
                            "type": "send_msg",
                            "quick_replies": quick_replies,
                            "uuid": uuidv4()
                        }
                    ],
                    "exits": exits
                };
                row.NodeUUID = nodeId;
                row.rapidProNode = sendMessageNode;
                nodesById[nodeId] = sendMessageNode;
                flow.nodes.push(sendMessageNode);
            }
        }
    }
}