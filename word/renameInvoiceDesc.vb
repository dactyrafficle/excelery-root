Option Explicit

Sub renameInvoiceDescriptions()

  Dim tbl As Table
  Set tbl = ActiveDocument.Sections(1).Range.Tables(1)
  tbl.Select
  
  Dim n_rows As Long, n_cols As Long
  n_rows = tbl.Rows.Count
  n_cols = tbl.Columns.Count

  Dim y As Long
  For y = 1 To n_rows
     
    Dim r As Range
    Set r = tbl.Cell(y, 1).Range
    r.End = r.End - 1 'excise le spleen

    Dim master As Variant
    master = Right(Trim(r.Text), 3) + 0

    tbl.Cell(y, 2).Range.Text = getDescFromMaster(master)

  Next y

End Sub

 

Function getDescFromMaster(master As Variant)

  Dim temp As Variant
  
  Select Case master
  
  Case 543: temp = "Sugar Free Peanut Butter Wafer Cookies, 255g"
  Case 521: temp = "Sugar Free Lemon Wafer Cookies, 255g"
  Case 525: temp = "Sugar Free Vanilla Wafer Cookies, 255g"
  Case 526: temp = "Sugar Free Chocolate Wafer Cookies, 255g"
  Case 527: temp = "Sugar Free Strawberry Wafer Cookies, 255g"
  Case 524: temp = "Sugar Free Coconut Wafer Cookies  , 255g"
  Case 558: temp = "Sugar Free Key Lime Wafer Cookies, 255g"
  Case 519: temp = "Sugar Free Orange Creme Wafers Cookies, 255g"
  Case 564: temp = "Sugar Free Iced Almonette Cookies, 227g"
  Case 551: temp = "Sugar Free Shortbread Cookies, 227g"
  Case 549: temp = "Sugar Free Fudge Brownie Chocolate Chip Cookies, 227g"
  Case 550: temp = "Sugar Free Chocolate Chip Cookies, 227g"
  Case 552: temp = "Sugar Free Oatmeal Cookies, 227g"
  Case 566: temp = "Sugar Free Pecan Chocolate Chip Cookies, 227g"
  Case 565: temp = "Sugar Free Pecan Shortbread Cookies, 227g"
  Case 557: temp = "Sugar Free Coconut Cookies, 200g"
  Case 559: temp = "Sugar Free Peanut Butter Cookies, 227g"
  Case 568: temp = "Sugar Free Fudge Striped Shortbread Cookies, 320g"
  Case 724: temp = "No Sugar Added Chocolate Chip Cookies, 480g"
  Case 725: temp = "No Sugar Added Oatmeal Cookies, 480g"
  Case 70: temp = "Sugar Free Apple Cinnamon Breakfast Cookies, 200g"
  Case 71: temp = "Sugar Free Oatmeal Blueberry Breakfast Cookies, 200g"
  Case 72: temp = "Sugar Free Chocolate Banana Breakfast Cookies, 200g"
  Case 283: temp = "Chocolate Chip Cookies, 200g"
  Case 208: temp = "Chocolate Chip Cookies, 350g"
  Case 635: temp = "Coconut Cookies, 420g"
  Case 692: temp = "Coconut Cookies, 630g"
  Case 285: temp = "Coconut Cookies, 200g"
  Case 210: temp = "Coconut Cookies, 350g"
  Case 278: temp = "Lemon Coconut Cookies, 200g"
  Case 229: temp = "Lemon Coconut Cookies, 350g"
  Case 228: temp = "Coconut Dark Chocolate Cookies, 350g"
  Case 637: temp = "Oatmeal Cookies, 420g"
  Case 690: temp = "Oatmeal Cookies, 630g"
  Case 603: temp = "Sugar Free Oatmeal Cookies, 720g"
  Case 277: temp = "Oatmeal Cookies, 200g"
  Case 227: temp = "Oatmeal Cookies, 350g"
  Case 284: temp = "Oatmeal Raisin Cookies, 200g"
  Case 215: temp = "Oatmeal Raisin Cookies, 350g"
  Case 226: temp = "Oatmeal Chocolate Chip Cookies, 350g"
  Case 219: temp = "Fudge Striped Oatmeal Cookies, 350g"
  Case 288: temp = "Oatmeal Cranberry Flaxseed Cookies, 200g"
  Case 256: temp = "Oatmeal Cranberry Flaxseed Cookies, 350g"
  Case 255: temp = "Oatmeal Dark Chocolate Flaxseed Cookies, 350g"
  Case 230: temp = "Fudge Striped Almonette Cookies, 350g"
  Case 231: temp = "Almond Delight Cookies, 350g"
  Case 232: temp = "Almond Crunch Cookies, 350g"
  Case 260: temp = "Vanilla Wafer Cookies, 241g"
  Case 430: temp = "Chocolate Wafer Cookies, 241g"
  Case 261: temp = "Strawberry Banana Wafer Cookies, 241g"
  Case 431: temp = "Coconut Creme Wafer Cookies, 241g"
  Case 300: temp = "Fudge Coated Vanilla Wafer Cookies, 235g"
  Case 301: temp = "Fudge Coated Chocolate Wafer Cookies, 235g"
  Case 302: temp = "Fudge Coated Peanut Butter Wafer Cookies, 235g"
  Case 303: temp = "Fudge Coated Salted Caramel Wafer Cookies, 235g"
  Case 799: temp = "Vanilla Wafer Cookies, 400g"
  Case 512: temp = "Chocolate Wafer Cookies, 300g"
  Case 513: temp = "Strawberry Wafer Cookies, 300g"
  Case 577: temp = "Coconut Creme Wafer Cookies, 300g"
  Case 516: temp = "Strawberry Banana Wafer Cookies, 300g"
  Case 632: temp = "Raspberry Wafer Cookies, 300g"
  Case 582: temp = "Chocolate & Raspberry Wafer Cookies, 300g"
  Case 584: temp = "Chocolate & Caramel Wafer Cookies, 300g"
  Case 574: temp = "Key Lime Wafer Cookies, 300g"
  Case 573: temp = "Banana Wafer Cookies, 300g"
  Case 576: temp = "Peanut Butter Wafer Cookies, 300g"
  Case 510: temp = "Lemon Wafer Cookies, 300g"
  Case 349: temp = "Iced Almonette Cookies, 250g"
  Case 360: temp = "Vanilla Shortbread Cookies, 250g"
  Case 358: temp = "Assorted Festive Cookies, 300g"
  Case 357: temp = "Gingerbread Cookies, 300g"
  Case 361: temp = "Snickerdoodle Cookies, 300g"
  Case 363: temp = "Chocolate Mint Wafer Cookies, 300g"
  Case 588: temp = "Gingerbread Wafer Cookies, 300g"
  Case 589: temp = "White Chocolate Cranberry Wafer Cookies, 300g"
  Case 741: temp = "Holiday Treats Cookies, 400g"
  Case 743: temp = "Holiday Gingerbread Cookies, 400g"
  Case 750: temp = "Vanilla Wafer Cookies, 432g"
  Case 751: temp = "Banana Wafer Cookies, 432g"
  Case 235: temp = "Smores Cookies, 240g"
  Case 587: temp = "Smores Wafer Cookies, 300g"
  Case 504: temp = "Pumpkin Spice Wafer Cookies, 300g"
  Case 500: temp = "Apple Crisp Wafer Cookies, 300g"
  Case 234: temp = "Pumpkin Spice Cookies, 300g"
  Case 667: temp = "Key Lime Wafer Cookies, 400g"
  Case 668: temp = "Coconut Creme Wafer Cookies, 400g"
  Case 669: temp = "Peanut Butter Wafer Cookies, 400g"
  Case 797: temp = "Strawberry Wafer Cookies, 400g"
  Case 798: temp = "Chocolate Wafer Cookies, 400g"
  Case 663: temp = "Blueberry Wafer Cookies, 400g"
  Case 664: temp = "Cinnamon Bun Wafer Cookies, 400g"
  Case 9564: temp = "Roll Pack Wing Display, 3182g"
  Case 9280: temp = "Wafer Wing Displays, 3182g"
  Case 9563: temp = "Bakery Sign Set Universal Rack, 454g"
  Case 9022: temp = "Universal Wire Rack, 12272g"
  Case 9381: temp = "Christmas Wing Display, 3182g"

  Case Else: temp = "CORRECT"
  
  End Select

  getDescFromMaster = temp

End Function