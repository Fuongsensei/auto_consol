
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    
  }
});
const btnAddHold : HTMLButtonElement = document.querySelector(".btnAddHold") as HTMLButtonElement;
const lbelsHold : HTMLInputElement = document.querySelector(".lbHoldNumbers") as HTMLInputElement;
const parrent : HTMLDivElement = document.querySelector(".containInputAndButton") as HTMLDivElement;

const sleep = (ms:number) => new Promise( (res,rej) => {
if (ms>0){
  setTimeout(res,ms);
}
else{
  rej("Time is invalid !")
}
} 
)


export  function run() {
  try {
    Excel.run(async (context) => {

      const range = context.workbook.getSelectedRange();

  
      range.load("address");
      range.load("values");



    context.sync();
      for(let i = 0 ;i<range.values.length;i++){
        for(let j = 0 ; j<range.values[i].length;j++)
          {
            lbelsHold.textContent= range.values[i][j];
          
          console.log(range.values[i][j]);
           sleep(1000);
        }
      }
    });

  } catch (error) {
    console.error(error);
  }
}


/*btnAddHold.addEventListener('click',  () => {
    btnAddHold.disabled = true; 
    
    run(); 
    
    btnAddHold.disabled = false; 
});*/

btnAddHold.addEventListener('click',async () => {
    console.log("Bắt đầu hành hạ CPU...");
    

    const start = Date.now();
    while (Date.now() - start < 5000) {
        await sleep(100);
    }
    
    console.log("Hết 5 giây, UI mới sống lại!");
});