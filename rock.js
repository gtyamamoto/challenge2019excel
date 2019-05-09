


const XLSX = require('xlsx');

const json2xls = require('node-xlsx');
const path = require('path');
const fs = require('fs');
const _ = require('lodash');

async function main(){
    const startTime = new Date().getTime();
    console.log('solution started');
    const fileMovPath = path.join(__dirname, `MovtoITEM.xlsx`);
    const fileSalPath = path.join(__dirname, `SaldoITEM.xlsx`);
    const workbookMov = XLSX.readFile(fileMovPath, {
        cellDates: true
    });
    const workbookSal = XLSX.readFile(fileSalPath, {
        cellDates: true
    });
    var sheet_name_listMov = workbookMov.SheetNames;
    const SheetMov = XLSX.utils.sheet_to_json(workbookMov.Sheets[sheet_name_listMov[0]]);
    var sheet_name_listSal = workbookSal.SheetNames;
    const SheetSal = XLSX.utils.sheet_to_json(workbookSal.Sheets[sheet_name_listSal[0]]);
    const itemsID = SheetSal.map(el=>el.item);
    const datesID = [...new Set(SheetMov.map(el => el.data_lancamento.toString()))].map(el=>new Date(el)).sort((a,b)=>a-b);
    
  let data = [];
   for(let date of datesID){
      for(let item of itemsID){
          const itemsFromDate = SheetMov.filter(el=>el.data_lancamento.getTime()===date.getTime()&&el.item===item);
          if(itemsFromDate.length){
            const findFirstRegistry = _.findLast(data,el=>el.item==item);
            let itemRef = SheetSal.find(el=>el.item==item);
            let firstQty = itemRef.qtd_inicio;
            let firstVal = itemRef.valor_inicio;
            if(findFirstRegistry){
                firstQty = findFirstRegistry.quantidade;
                firstVal = findFirstRegistry.valor;
            }
            const qtyIn = _.sumBy(itemsFromDate.filter(el=>el.tipo_movimento==='Ent'),'quantidade');
            const qtyOut = _.sumBy(itemsFromDate.filter(el=>el.tipo_movimento==='Sai'),'quantidade');
            const valIn = _.sumBy(itemsFromDate.filter(el=>el.tipo_movimento==='Ent'),'valor');
            const valOut =_.sumBy(itemsFromDate.filter(el=>el.tipo_movimento==='Sai'),'valor');
            let finalQty = firstQty + qtyIn   - qtyOut;
            let finalVal = firstVal  + valIn - valOut;
            const obj = {
                item,
                data_lancamento: date,
                entrada_quantidade: qtyIn,
                entrada_valor:valIn,
                saida_quantidade:qtyOut,
                saida_valor:valOut,
                qtd_inicio:firstQty,
                valor_inicio:firstVal,
                quantidade:finalQty,
                valor:finalVal
            }
        
            data.push(obj);

          }
        
      } 
   }

  
   const xlsxbuild = json2xls.build([{name:'balanco_diario',data:[Object.keys(data),...data.map( Object.values )]}])
   fs.writeFileSync('data.xlsx', xlsxbuild, 'binary');
  console.log(`time to end: ${(new Date().getTime()-startTime)/1000} seconds`)
}
main()