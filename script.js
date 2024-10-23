document.getElementById('excelUpload').addEventListener('change', handleFileUpload);
document.getElementById('printBills').addEventListener('click', () => window.print());

function handleFileUpload(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    displayBills(jsonData);
  };

  reader.readAsArrayBuffer(file);
}
function numberToWords(num) {
    if (num === 0) return "zero";

    const belowTwenty = [
        "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
        "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"
    ];
    const tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];
    const thousands = ["", "thousand", "million", "billion"];

    function convertBelowThousand(n) {
        if (n < 20) return belowTwenty[n];
        else if (n < 100) return tens[Math.floor(n / 10)] + (n % 10 !== 0 ? " " + belowTwenty[n % 10] : "");
        else return belowTwenty[Math.floor(n / 100)] + " hundred" + (n % 100 !== 0 ? " " + convertBelowThousand(n % 100) : "");
    }

    let result = "";
    let thousandCounter = 0;

    while (num > 0) {
        if (num % 1000 !== 0) {
            result = convertBelowThousand(num % 1000) + (thousands[thousandCounter] ? " " + thousands[thousandCounter] : "") + " " + result;
        }
        num = Math.floor(num / 1000);
        thousandCounter++;
    }

    return result.trim();
}

function determineBranchType(branchName) {
  let branchType = 'UBI'; // Default to UBI if nothing else matches

  // Convert branchName to lowercase for case-insensitive comparison
  const lowerBranchName = branchName.toLowerCase();

  if (lowerBranchName.includes('ecb')) {
      branchType = 'eCB';
      branchName = branchName.replace(/eCB/gi, '').trim(); // Remove 'eCB' (case-insensitive)
  } else if (lowerBranchName.includes('eab')) {
      branchType = 'eAB';
      branchName = branchName.replace(/eAB/gi, '').trim(); // Remove 'eAB' (case-insensitive)
  } else if (lowerBranchName.includes('ubi')) {
      branchType = 'UBI';
      branchName = branchName.replace(/UBI/gi, '').trim(); // Remove 'UBI' (case-insensitive)
  }

  return { branchName, branchType };
}


function displayBills(data) {
    const container = document.getElementById('billContainer');
    container.innerHTML = ''; 
    const apptitle = document.querySelector('.app-title');
    const excelUpload = document.querySelector('#excelUpload');
    const labelUpload = document.querySelector('.label-upload');
    apptitle.style.display='none';
    excelUpload.style.display='none';
    labelUpload.style.display='none';
    container.style.boxShadow="0";
    const headers = data[0]; 
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const { branchName, branchType } = determineBranchType(row[headers.indexOf('Branch')]); 
      const customerName = row[headers.indexOf('Customer Name')];  
      // Check if customerName contains '-' and extract cities
        let city1 = '';
        let city2 = '';
        let finalCustomerName = customerName;  // This will hold the modified customer name without cities

        if (customerName.includes(' - ')) {
            // Split customer name at ' - ' to separate the cities part
            [finalCustomerName, cities] = customerName.split(' - ');
            // Split cities by '&'
            const citiesArray = cities.split('&').map(city => city.trim());
            // Store cities in city1 and city2
            city1 = citiesArray[0] ? citiesArray[0] : '';
            city2 = citiesArray[1] ? citiesArray[1] : '';
        }
      const Amount = row[headers.indexOf('Amount')];
      // Check if amount contains '+' and extract charges
        let charge1 = '';
        let charge2 = '';
        let finalAmount = Amount;  // This will hold the modified amount after removing the charges

        if (String(Amount).includes('+')) {
            // Split the amount at '+'
            const chargesArray =Amount.split('+').map(charge => charge.trim());

            // Store charges in charge1 and charge2
            charge1 = chargesArray[0] ? chargesArray[0] : '';
            charge2 = chargesArray[1] ? chargesArray[1] : '';

            // Remove the charges part from the amount
            finalAmount = '';  // If both charges are extracted, the final amount field can be emptied
        }
      const totalAmount = row[headers.indexOf('Total')];
      const formattedTotalAmount = new Intl.NumberFormat('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(totalAmount);
      const totalAmountInWords = numberToWords(totalAmount);
      const bill = document.createElement('div');
      bill.classList.add('page');
      bill.innerHTML = `
        <div class="topper">
        <div style="display:inline;">
        <img src="logo.png" width="140px" height="140px">
        </div>
        <div style="display:inline;">
        <img src="brandname.png" height="110px">
        </div>
    </div>
    <div class="container-b">
        <div class="intro">
            <div class="cin"><h4>CIN :  U67190TZ2017PTC029738 </h4></div>
            <div class="report">
                    <h3 class="cpv">CONTACT POINT VERIFICATION / DUE DELIGENCE REPORT</h3> 
                            <p> Registered Office : No:  120, Radha Complex (No: 6),<br>
                           Karungalpatti Main Road, Salem ,  Tamilnadu– 636 006.<br>
                         Ph.No: 0427 – 2468093, E-Mail: truevalue.salem@gmail.com  <br>          
                                <div><h4 style="display: inline;">GST No :</h4><h4 style="display: inline; font-size: 16px;"> 33AAGCT5341Q1ZX</h4></div>
                            </p>
            </div>
        </div>
        <div class="invoice-title"><p style="font-size:20px; font-weight:bold;">TAX INVOICE</p></div>
        <div class="invoice-details">
            <table>
                <tr>
                    <td style="width: 50px !important;"><p style="font-size:20px;">Invoice No:</p></td>
                    <td style="padding-top: 20px;"><span style="font-size:20px;">${row[headers.indexOf('Bill No')]}</span></td>
                    <td style="padding-left: 60px !important; width:70px;"><p style="font-size:20px;">Date&nbsp;:</p></td>
                    <td><span>${row[headers.indexOf('Bill Date')]}</span></td>
                </tr>
                <tr>
                    <td style="width: 90px !important; "><p style="font-size:20px;">Bill To:</p></td>
                    <td style="padding-top:10px ;"><span style="font-size:20px;">Union Bank of India</span></td>
                    <td style="padding-left: 50px !important;"><p style="width: 100px !important; font-size:20px;">Branch :</p></td>
                    <td><span style="font-size:20px; text-align:start;">${branchName}</span></td>
                </tr>
                <tr>
                    <td style="width: 90px !important;"><p style="font-size:20px;">GST:</p></td>
                    <td style="padding-top:10px;"><span style="font-size:20px;">33AAACU0564G3ZM</span></td>
                    <td></td>
                    <td ><span style="font-size:20px;">${branchType}</span></td>
                </tr>
            </table>
        </div>
        <div class="case-details">
            <table>
                <thead style="border-top: 3px solid black;border-bottom: 3px solid black;" >
                    <tr style="font-size: 20px;">
                        <th style="width: 10% !important; border-right:  3px solid black; ">S.No</th>
                        <th style="width: 40% !important; border-right:  3px solid black;">Customer Name</th> 
                        <th style="width:13%; border-right:3px solid black;">Verified For</th>
                        <th style="width:12%; border-right:3px solid black;">Amount</th>
                        <th style="width:10%; border-right:3px solid black;">GST</th>
                        <th style="width:15%;">Total</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="width: 10%; border-right:3px solid black; font-size: 20px; font-weight: bold; text-align: center;">1</td>
                        <td style="width: 40% !important; border-right:  3px solid black; font-size: 20px; font-weight: bold; padding-left:5px !important;">${finalCustomerName}</td>
                        <td style="width:13%; border-right:3px solid black; font-size: 20px; text-align: center; font-weight:bold;">${row[headers.indexOf('Type of Loan')]}</td>
                        <td style="width:12%; border-right:3px solid black; font-size: 20px; text-align: center;"> ${finalAmount|| ''}</td>
                        <td style="width:10%; border-right:3px solid black; font-size: 20px; text-align: center;"> ${row[headers.indexOf('GST')]|| ''}</td>
                        <td style="width:15%; font-size: 20px; text-align: center;"> ${totalAmount|| ''}</td>
                    </tr>
                    <tr>
                    <td style="width: 10%; border-right:3px solid black; font-size: 20px; font-weight: bold; text-align: center;"></td>
                    <td style="width: 40% !important; border-right: 3px solid black; font-size: 20px; font-weight: bold; padding-left:5px !important;">${city1 ?? ''}</td>
                    <td style="width:13%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:12%; border-right:3px solid black; font-size: 20px; text-align: center;">${charge1 ?? ''}</td>
                    <td style="width:10%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:15%; font-size: 20px; text-align: center; padding-left:5px !important;"></td>
                     </tr>
                    <tr>
                       <td style="width: 10%; border-right:3px solid black; font-size: 20px; font-weight: bold; text-align: center;"></td>
                    <td style="width: 40% !important; border-right: 3px solid black; font-size: 20px; font-weight: bold; padding-left:5px !important;">${city2 ?? ''}</td>
                    <td style="width:13%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:12%; border-right:3px solid black; font-size: 20px; text-align: center;">${charge2 ?? ''}</td>
                    <td style="width:10%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:15%;  font-size: 20px; text-align: center;"></td>
                    </tr>
                    <tr style="height:80px;">
                    <td style="width: 10%; border-right:3px solid black; font-size: 20px; font-weight: bold; text-align: center;"></td>
                    <td style="width: 40% !important; border-right: 3px solid black; font-size: 20px; font-weight: bold;"></td>
                    <td style="width:13%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:12%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:10%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                    <td style="width:15%;  font-size: 20px; text-align: center;"></td>
                    </tr>
                    <tr style="border-top: 3px solid black;">
                        <td style="width: 10%; border-right:3px solid black;"></td>
                        <td style="width: 40%; border-right:3px solid black;"></td>
                        <td style="width:13%; border-right:3px solid black; font-size: 20px; text-align: center;"></td>
                        <td style="width:12%; border-right:3px solid black; font-size: 20px; text-align: center; font-weight:bold;">${finalAmount}</td>
                        <td style="width:10%; border-right:3px solid black; font-size: 20px; text-align: center; font-weight:bold;">${row[headers.indexOf('GST')]}</td>
                        <td style="width:15%;  font-size: 20px;  text-align:center; font-weight:bold;">${totalAmount}</td>
                    </tr>
                    <tr style="border-top: 3px solid black;" id="nested-data">
                        <td colspan="4" style="text-align: center; font-size: 21px; font-weight: bold; border-right: 3px solid black;" id="nested-data">Total Cost</td>
                        <td colspan="2" style="text-align: center; font-size: 21px; font-weight: bold;" id="nested-data"> &#8377;${formattedTotalAmount}</td>
                    </tr>
                    <tr style="border-top: 3px solid black;" id="nested-data">
                      <!--   <td id="nested-data" style="font-size: 20px; font-weight: bold;padding-left: 5px !important;">In words:</td>-->
                        <td colspan="6" style=" font-size: 18px; font-weight: bold; text-transform: capitalize;" id="nested-data">In Words: <span style="margin-left:30px !important; font-size:20px;"> Rupees ${totalAmountInWords} only </span></td>
                    </tr>    
                </tbody>
            </table>
        </div>
        <div class="tvpc-details" style="position: relative;">
            <p>Please Make Transfer in the Name</p>
            <p>M/s. True Value Professional  Consultancy  Pvt  Ltd,</p>
            <p>Bank&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;Union Bank of India - Salem Branch </p>
            <p>A/C No&nbsp;: 050011100003730</p>
            <span>IFSC&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: UBIN0900290</span>

            <!-- Bank seal image with absolute positioning -->
            <img class="seal" src="Digital-seal-img.png" width="140px" style="position: absolute; right: 20%; bottom:36%">
            <p style="text-align: end; padding-top:25px;">For True Value Professional Consultancy Pvt Ltd.,</p>
            <span>Place: Salem</span>

            <!-- Signature image with absolute positioning -->
            <img class="seal" src="sign.jpg" width="140px" style="position: absolute; right: 25%; bottom:10%">

            <div class="sign" style="padding-top:15px;">
                <p style="display: inline;">Tamilnadu</p>
                <p style="display: inline; padding-left: 250px;">Manager / Head of the Department</p>
            </div>
        </div>

        
    </div>
    
      `;
  
      container.appendChild(bill);
    }
  }

  document.getElementById('saveImages').addEventListener('click', saveAllImages);

function saveAllImages() {
  const bills = document.querySelectorAll('.page');

  bills.forEach((bill, index) => {
    const invoiceNo = bill.querySelector('span').textContent; 

    const clonedBill = bill.cloneNode(true);
    
    const wrapper = document.createElement('div');
    
    const padding = 20;
    wrapper.style.padding = `${padding}px`;
    wrapper.style.backgroundColor = '#fff'; 
    wrapper.style.width = `${bill.offsetWidth + padding * 2}px`;
    wrapper.style.display = 'inline-block'; 

    wrapper.appendChild(clonedBill);

    document.body.appendChild(wrapper);

    html2canvas(wrapper, {
      backgroundColor: '#fff', 
      scale: 2, 
      useCORS: true
    }).then(canvas => {
      const link = document.createElement('a');
      link.download = `${invoiceNo}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
      document.body.removeChild(wrapper);
    });
  });
}

  

  

  
