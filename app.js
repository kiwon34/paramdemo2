
$(function () {
    var colM = [
              { title: "ShipCountry", width: 120, dataIndx: "ShipCountry" },
              { title: "Customer Name", width: 130, dataIndx: "ContactName" },
              { title: "Order ID", width: 100, dataIndx: "OrderID" , dataType:"integer"},
              { title: 'Dates', styleHead: {}, colModel:[
                  { title: "Order Date", width: "100", dataIndx:"OrderDate", dataType:"date" },
                  { title: "Required Date", width: 100 , dataIndx:"RequiredDate", dataType:"date"},
                  { title: "Shipped Date", width: 100, dataIndx: "ShippedDate" }
              ]},
              { title: "Freight", width: 120, format: '$##,###.00',
                  summary: {
                      type: "sum"
                  },
                  dataType: "float", dataIndx: "Freight"
              },
              { title: "Shipping Via", width: 130, dataIndx: "ShipVia" },		    
              { title: "Shipping Name", width: 160, dataIndx:"ShipName" },
              { title: "Shipping Address", width: 300, dataIndx:"ShipAddress" },
              { title: "Shipping City", width: 100, dataIndx:"ShipCity" },
              { title: "Shipping Region", width: 100,dataIndx:"ShipRegion" },
              { title: "Shipping PostalCode", width: 100, dataIndx:"ShipPostalCode" }
          ];
          var dataModel = {
              location: "remote",
              dataType: "JSON",
              method: "GET",
              url: "https://paramquery.com/Content/orders.json"
              //url: "/pro/orders/get",//for ASP.NET
              //url: "orders.php",//for PHP
          };
          function exportData(format){
              var blob = this.exportData({
                  format: format
              }) 
              if(typeof blob === "string"){                            
                  blob = new Blob([blob]);
              }
              saveAs(blob, "pqGrid."+ format );
          }
  
          //returns menu items for header cells, can be a callback or an array
          function headItems(evt, ui){
              return [
                  {
                      name: 'Rename',
                      action: function (evt, ui) {
                          var grid = this,
                              column = ui.column,
                              title = column.title;
                          title = prompt("Enter new column name", title);
                          if (title) {
                              grid.Columns().alter(function () {
                                  column.title = title;
                              })
                          }
                      }
                  },
                  {
                      name:'Toggle filter row',
                      action: function(){
                          this.option('filterModel.header', !this.option('filterModel.header'));
                          this.refresh();
                      }
                  },
                  {
                      name: "color",
                      subItems:[
                          {
                              name: 'none',
                              disabled: !ui.column.styleHead.background,
                              action: function(evt, ui, item){
                                  delete ui.column.styleHead.background
                                  this.refreshHeader();
                              }
                          },
                          {
                              name:'green',
                              disabled: (ui.column.styleHead.background == 'lightgreen'),
                              action: function(evt, ui, item){
                                  ui.column.styleHead.background = 'lightgreen'
                                  this.refreshHeader();
                              }
                          },
                          {
                              name:'red',
                              disabled: (ui.column.styleHead.background == 'red'),
                              action: function(evt, ui, item){
                                  ui.column.styleHead.background = 'red'
                                  this.refreshHeader();
                              }
                          },
                          {
                              name:'blue',
                              disabled: (ui.column.styleHead.background == 'lightblue'),
                              action: function(evt, ui, item){
                                  ui.column.styleHead.background = 'lightblue'
                                  this.refreshHeader();
                              }
                          }
                      ]
                  }
              ]
          }
  
          //returns menu items for body cells, can be a callback or an array
          function bodyItems(evt, ui){
              return [
                  {
                      name: 'Add column',
                      action: function (evt, ui) {
                          var grid = this,
                              col = ui.column,
                              parent = col.parent,
                              CM = parent ? parent.colModel : grid.option('colModel'),
                              ci = CM.indexOf(col),
                              title = prompt("Please enter column name", "New column");
  
                          if (title) {
                              grid.Columns().add([{
                                  title: title,
                                  dataIndx: Math.random(),
                                  width: 150
                              }], ci, CM)
                          }
                      }
                  },
                  {
                      name: 'Delete column',
                      icon: 'ui-icon ui-icon-trash',
                      action: function (evt, ui) {
                          var grid = this,
                              col = ui.column,
                              CM = col.parent ? col.parent.colModel : grid.option('colModel'),
                              ci = CM.indexOf(col);
  
                          grid.Columns().remove(1, ci, CM)
                      }
                  },
                  {
                      name: 'Frozen columns',
                      subItems: [
                          (this.option('freezeCols')?{
                              name: 'none',                           
                              action: function(){
                                  this.option('freezeCols', 0);
                                  this.refresh();
                              }
                          }: null),
                          'separator',
                          {
                              name: '1',
                              disabled: (this.option('freezeCols')==1),
                              action: function(){
                                  this.option('freezeCols', 1);
                                  this.refresh();
                              }
                          },
                          {
                              name: '2',
                              disabled: (this.option('freezeCols')==2),
                              action: function(){
                                  this.option('freezeCols', 2);
                                  this.refresh();
                              }
                          }
                      ]
                  },                
                  {
                      name: 'Export',
                      subItems: [
                          {
                              name: 'csv',
                              action: function(){
                                  exportData.call(this, 'csv');                                   
                              }
                          },
                          {
                              name: 'html',
                              action: function(){
                                  exportData.call(this, 'html');
                              }
                          },
                          {
                              name: 'json',
                              action: function(){
                                  exportData.call(this, 'json');
                              }
                          },
                          {
                              name: 'xlsx',
                              action: function(){
                                  this.one("workbookReady", function(e, w){
                                    debugger;
                                    let sheet = w.workbook.sheets[0];
                                    sheet.frozenRows = 0;
                                    sheet.frozenCo1s = 0;
                                  });
                                  exportData.call(this, 'xlsx');
  
  
                              }
                          }
                      ]                        
                  },
                  'separator',
                  {
                      name: "Undo",
                      icon: 'ui-icon ui-icon-arrowrefresh-1-n',
                      disabled: !this.History().canUndo(), 
                      action: function(evt, ui){
                          //debugger;
                          this.History().undo();
                      }
                  },
                  {
                      name: "Redo",
                      icon: 'ui-icon ui-icon-arrowrefresh-1-s',
                      disabled: !this.History().canRedo(), 
                      action: function(evt, ui){
                          //debugger;
                          this.History().redo();
                      }
                  },
                  'separator',
                  {
                      name: "Copy",
                      icon: 'ui-icon ui-icon-copy',
                      shortcut: 'Ctrl - C',
                      tooltip: "Works only for copy / paste within the same grid",
                      action: function(){
                          // grid range copy to Excel (ctl+c)
                          this.copy();
                      }
                  },
                  {
                      name: "Paste",
                      icon: 'ui-icon ui-icon-clipboard',
                      shortcut: 'Ctrl - V',
                      //disabled: !this.canPaste(),
                      action: function(){                        
                          this.paste();
                          //this.clearPaste();
                      }
                  }
              ]
          }
          
          var obj = {
              height: 500,               
              dataModel: dataModel,            
              colModel: colM, 
              complete: function(){
                  this.flex();
              },            
              contextMenu: {
                  on: true,
  
                  //header context menu items.
                  headItems: headItems,
  
                  //body context menu items
                  cellItems: bodyItems,                
  
                  //image context menu items
                  imgItems: [
                      {
                          name: 'Delete',
                          action: function (evt, ui) {
                              var Pic = this.Pic()
                              Pic.remove( Pic.getId( ui.ele ) )
                          }
                      }
                  ]
              },
              columnTemplate: {
                  styleHead: {}
              },           
              showTitle: false,
              menuIcon: true
          };
          pq.grid("#context_menu", obj);
  
  })
  