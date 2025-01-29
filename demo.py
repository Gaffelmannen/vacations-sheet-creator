#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import pygsheets 
    
def read_title():
    client = pygsheets.authorize(
        service_account_file="vacation-lists-aecd2c47a267.json"
    )
    spreadsht = client.open("vacations-sheet") 
 
    worksht = spreadsht.worksheet("title", "Sheet1") 
    
    worksht.cell("A1").set_text_format("bold", True).value = "Item"
    
    worksht.update_values("A2:A6", 
                                    [
                                        ["Mouse"], ["Keyboard"],  
                                        ["Computer"], ["Monitor"],  
                                        ["Headphones"]
                                    ]) 
    
    worksht.cell("B1").set_text_format("bold", True).value = "Price"
    worksht.update_values("B2:B6", [[70.0], [100.0], [2000.0], [1500.0], [100.0]]) 
    
    worksht.add_chart(("A2", "A6"), [("B2", "B6")], "Webshop")

if __name__ == "__main__":
    read_title()