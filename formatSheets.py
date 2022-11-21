import random
from sheetFunctions import sheetFunctions
import datetime

class formatSheets:

    current = datetime.datetime.now()
    current_time = str(current.year) + "/" + str(current.month) + "/" + str(current.day) + " " + str(current.hour) + ":" + str(current.minute) + ":" + str(current.second)
    userLength = len(sheetFunctions.users)
    protectedSheetLength = len(sheetFunctions.protectedSheets)
    protectedRangeLength = len(sheetFunctions.protectedSheets.values())
    
    print("Detected "+  str(userLength) +" users with permissions set.")
    
    sheetsRange  = []
    for i in sheetFunctions.protectedSheetsId:
        sheetsRange.append(len(sheetFunctions.protectedSheets[i]))

    sheetsName = list(sheetFunctions.protectedSheetsId.keys())

    enteredSheets = {
        "requests": 
        [
            {
                "updateCells": {
                    "range":{
                        "sheetId": 0,
                    },
                    "fields": '*',
                }
            },
            {
                "unmergeCells": {
                    "range": {
                        "sheetId": 0,
                    }
                }
            },
            {
                "mergeCells": {
                    "mergeType": "MERGE_ROWS",
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 0,
                        "endRowIndex": 5,
                        "startColumnIndex": 0,
                    }
                }
            },
            {
                "mergeCells": {
                    "mergeType": "MERGE_ROWS",
                    "range": {
                    "sheetId": 0, 
                    "startRowIndex": 5,
                    "endRowIndex": 6,
                    "startColumnIndex": 1,
                    },
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredValue": {
                            "stringValue": "Total detected sheets = %s"%(sheetFunctions.sheetNumbers)
                        }
                    },
                    "fields": "userEnteredValue.stringValue"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredValue": {
                            "stringValue": "Total protected sheets = %s"%(protectedSheetLength)
                        }
                    },
                    "fields": "userEnteredValue.stringValue"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 2,
                        "endRowIndex": 3,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredValue": {
                            "stringValue": "Total emails = %s"%(userLength)
                        }
                    },
                    "fields": "userEnteredValue.stringValue"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 3,
                        "endRowIndex": 4,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredValue": {
                            "stringValue": "Date last run = %s"%(current_time)
                        }
                    },
                    "fields": "userEnteredValue.stringValue"
                }
            },
            {
                "repeatCell": {
                    "cell": {
                        "userEnteredValue": {
                            "stringValue": 'SHEET NAME'
                        },
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontSize": 16,
                                "bold": True}
                        }
                    },
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 5,
                        "endRowIndex": 6,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "fields": "userEnteredValue.stringValue, userEnteredFormat.textFormat"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 5,
                        "endRowIndex": 6,
                        "startColumnIndex": 1,
                        "endColumnIndex": 2
                    },
                    "cell": {
                        "userEnteredValue": {
                            "stringValue": "NAME / EMAIL"
                        },
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontSize": 16,
                                "bold": True}
                        }
                    },
                    "fields": "userEnteredValue.stringValue, userEnteredFormat.textFormat"
                }
            },
            {
                "repeatCell": {#VERTICAL FORMAT
                    "range": {
                        "sheetId": 0, 
                        "startRowIndex": 5,
                        "endRowIndex": 8,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 50,
                                "green": 50,
                                "blue": 50
                            },
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            },
            {
                "repeatCell": {#HORIZONTAL FORMAT
                    "range": {
                        "sheetId": 0,  
                        "startRowIndex": 5,
                        "endRowIndex": 8,
                        "startColumnIndex": 0,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 50,
                                "green": 50,
                                "blue": 50
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": 0,
                        "dimension": "COLUMNS"
                    }
                }
            }
        ]
    }

    increment = 0

    for cell in range(len(sheetsRange)):
        start = 8 + increment
        mergeRequests = {
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {
                    "sheetId": 0,
                    "startRowIndex": start,
                    "endRowIndex": start + sheetsRange[cell] + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
            },
        }
        textRequests = {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": start,
                    "endRowIndex": start + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredValue": {
                        "stringValue": sheetsName[cell]
                    },
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True}
                    }
                }, 
                "fields": "userEnteredValue.stringValue, userEnteredFormat.horizontalAlignment, userEnteredFormat.verticalAlignment, userEnteredFormat.textFormat"
            }
        }

        r = random.randint(15,80)
        b = random.randint(15,80)
        g = random.randint(15,80)


        merge_formatRequest = {
            "repeatCell": {
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": r,
                            "green": g,
                            "blue": b
                        }
                    }
                },
                "range": {
                    "sheetId": 0,
                    "startRowIndex": start,
                    "endRowIndex": start + sheetsRange[cell] + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }

        formatRequest = {
            "repeatCell": {
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": r ,
                            "green": g,
                            "blue": b
                        }
                    }
                },
                "range": {
                    "sheetId": 0,
                    "startRowIndex": start,
                    "endRowIndex": start + sheetsRange[cell] + 1,
                    "startColumnIndex": 1
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }
        enteredSheets['requests'].append(mergeRequests)
        enteredSheets['requests'].append(textRequests)
        enteredSheets['requests'].append(merge_formatRequest)
        enteredSheets['requests'].append(formatRequest)
        increment = increment + sheetsRange[cell] + 1

    clearAll = {
        "requests": [
            {
                "updateCells": {
                    "range":{
                        "sheetId": 0,
                    },
                    "fields": '*',
                }
            },
            {
                "unmergeCells": {
                    "range": {
                        "sheetId": 0,
                    }
                }
            }   
        ]
    }

    adjustColumnWidth = {
        "requests": [
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": 0,
                        "dimension": "COLUMNS"
                    }
                }
            }
        ]
    }
