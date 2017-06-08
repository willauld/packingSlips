package main

import (
	"bufio"
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"

	//"pflags"
	ole "github.com/go-ole/go-ole"
	"github.com/spf13/pflag"

	"github.com/noypi/xlsx"
	//"github.com/willauld/xlsx"
	//"github.com/xlsx"
)

var version = struct {
	major int
	minor int
}{1, 1}
var storeItems = map[string]string{
	"Full Ingredient Sake Kit": "Rice milled to ~60% 10 lbs.\nKoji 40 Oz.\nYeast #9\nLactic Acid 2 fl. Oz.\nYeast Nutrient 1 Oz.\nSpeedy Bentonite 2 Oz.",
	"Sake Ingredient Kit":      "Rice milled to ~60% 10 lbs.\nKoji 40 Oz.\nYeast #9",
	"Rice milled for Sake":     "Medium grain rice\nMilled to ~60% (Ginjo Level)\n10 lbs. bag",
	"Koji":                     "Rice milled to ~60% cultured with koji kin\n40 oz package",
	"Yeast #9":                 "Wyeast 4134 Sake",
	"Lactic Acid 88%":          "Lactic Acid 88%\n2Fl. Oz.",
	"Yeast Nutrient":           "Thiamin, vitamin B complex\n1 Oz.",
	"Speedy Bentonite":         "Bentonite clay wine clarifier\n2 Oz.",
	"Koji-kin":                 "15g Powdered Rice Koji Starter\nEnough to make 2 batches of 2.5 lbs. koji each\nAspergillus oryzae and rice flour\nPrinted Instructions",
	"Yeast #7":                 "White Labs WLP705",
	"Special Ginjo Koji-kin":   "2 x 1g Powdered Akita Konno Special Ginjo Koji Starter\nEach 1g packet makes 3.14 lbs koji (6.28 lbs. total)\nAspergillus oryzae\nPrinted Instructions",
}

type address struct {
	firstName   string
	lastName    string
	street      string
	city        string
	state       string
	country     string
	zipCode     string
	email       string
	phoneNumber string
}
type item struct {
	quantity       int
	title          string
	itemPrice      float32
	totalItemPrice float32
}
type order struct {
	billing      address
	shipping     address
	items        []item
	totalItem    float32
	shippingCost float32
	totalOrder   float32
}

func intMax(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// getQuantity gets the first integer in the string
func getQuantity(line string) int {
	var digit int
	quantity := 0
	firstDigit := false

	for i := 0; i < len(line); i++ {
		if line[i] >= '0' && line[i] <= '9' {
			firstDigit = true
			digit = int(line[i]) - int('0')
			quantity = quantity*10 + digit
		} else if firstDigit == true {
			return quantity
		}
	}
	return 0
}

// getTitle gets the first substring starting with [AZaz] to either a '$' or '--' sequence
func getTitle(line string) string {

	startChar := 0
	endChar := 0
	for i := 0; i < len(line); i++ {
		if line[i] >= 'A' && line[i] <= 'Z' || line[i] >= 'a' && line[i] <= 'z' {
			startChar = i
			break
		}
	}
	for i := 0; i < len(line); i++ {
		if line[i] == '-' && line[i+1] == '-' {
			endChar = i - 1
			break
		}
		if line[i] == '$' {
			endChar = i - 1
			break
		}
	}
	if endChar == 0 {
		endChar = len(line) - 1
	}
	for j := endChar; j >= 0; j-- {
		if line[j] != ' ' && line[j] != '\t' {
			endChar = j + 1
			break
		}
	}
	return line[startChar:endChar]
}

// getPrice returns the first float number passed the '$'
func getPrice(line string) float32 {

	var startChar int
	var price float32
	pastPoint := 0
	foundPoint := false
	price = 0.0

	for i := 0; i < len(line); i++ {
		if line[i] == '$' {
			startChar = i + 1
			break
		}
	}
	for j := startChar; j < len(line); j++ {

		if line[j] >= '0' && line[j] <= '9' {
			price = price*10.0 + float32(int(line[j])-int('0'))
			if foundPoint == true {
				pastPoint++
			}
		} else if line[j] == '.' {
			foundPoint = true
		}
	}
	price = price / float32(math.Pow(10.0, float64(pastPoint)))
	return price
}

// outputSpreadsheet updates the template spreadsheet and prints it
func outputSpreadsheet(purchase order, xlsxPath string, save bool) {

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	var err error
	defer func() {
		if nil != err {
			fmt.Println("err=", err)
		}
	}()

	excel, err := xlsx.CreateObject()
	if nil != err {
		fmt.Println(err)
		return
	}
	defer excel.Release()

	//xlsxPath := "C:\\home\\auld\\goDev\\src\\packingSlips\\packingSlipTemplate.xlsx"

	workbooks := excel.Workbooks()
	//workbook := workbooks.Create()
	workbook := workbooks.Open(xlsxPath)
	defer workbook.Close()

	sheet1 := workbook.Worksheets(1)

	//cell := sheet1.Range("a1")
	//a1Val := cell.ToString()
	//fmt.Println("a1Val: ", a1Val)
	//cell.PutValue("adrian guwapo")

	shipping := &purchase.shipping
	billing := &purchase.billing

	cell := sheet1.Range("c12")
	cell.PutValue(shipping.firstName + " " + shipping.lastName)

	cell = sheet1.Range("c13")
	cell.PutValue(shipping.street)

	cell = sheet1.Range("c14")
	cell.PutValue(shipping.city + ", " + shipping.state + " " + shipping.zipCode)

	cell = sheet1.Range("c15")
	cell.PutValue(shipping.country)

	cell = sheet1.Range("f12")
	cell.PutValue(billing.firstName + " " + billing.lastName)

	cell = sheet1.Range("f13")
	cell.PutValue(billing.street)

	cell = sheet1.Range("f14")
	cell.PutValue(billing.city + ", " + billing.state + " " + billing.zipCode)

	cell = sheet1.Range("f15")
	cell.PutValue(billing.country)

	cell = sheet1.Range("f16")
	cell.PutValue(billing.email)

	startRow := 19
	for i, v := range purchase.items {
		row := startRow + i

		loc := fmt.Sprintf("b%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(v.title)

		loc = fmt.Sprintf("c%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(storeItems[v.title])

		loc = fmt.Sprintf("e%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(v.quantity)

		loc = fmt.Sprintf("f%d", row)
		sheet1.Range(loc).PutValue(v.itemPrice) //&&&&&&&&&
		//cell = sheet1.Range(loc)
		//cell.PutValue(v.itemPrice)

		nlCount := strings.Count(v.title, "\n")
		nlCount = intMax(nlCount, strings.Count(storeItems[v.title], "\n"))
		rowRange := fmt.Sprintf("%d:%d", row, row)
		sheet1.Range(rowRange).PutRowHeight(12.4 * float64(nlCount+1))
	}
	loc := fmt.Sprintf("g%d", 28)

	cell = sheet1.Range(loc)
	cell.PutValue(purchase.shippingCost)

	if save == true {
		saveTo := "C:\\home\\auld\\personal\\Sake'\\PackingSlips\\email\\" +
			billing.firstName + ".xlsx"
		fmt.Printf("Saveing Excel to: %s\n", saveTo)
		workbook.Save(saveTo)
	} else {
		sheet1.PrintOut(1, 1, 1) // params: fromPage, toPage, copies
	}
	//os.Exit(1)

	//filepath2 := "c:\\temp\\a.xlsx"
	//workbook.Save(filepath2)
}

// printPurchaseRecord displays the purchase information on the terminal
func printPurchaseRecord(purchase order, i int) {
	b := purchase.billing
	s := purchase.shipping
	fmt.Println("")
	fmt.Printf("%d: Shipping:                            Billing:\n", i)
	fmt.Printf("   =========                            ========\n")
	s1 := s.firstName + " " + s.lastName
	s2 := b.firstName + " " + b.lastName

	fmt.Printf("%d: %-35s  %-35s\n", i, s1, s2)
	fmt.Printf("%d: %-35s  %-35s\n", i, s.street, b.street)
	// cancat first and last and then print in fixed space
	l1 := s.city + ", " + s.state + " " + s.zipCode
	l2 := b.city + ", " + b.state + " " + b.zipCode
	fmt.Printf("%d: %-35s  %-35s\n", i, l1, l2)
	fmt.Printf("%d: %-35s  %-35s\n", i, s.country, b.country)
	fmt.Printf("%d: %-35s  %-35s\n", i, b.phoneNumber, b.email)
	fmt.Println("")
	//fmt.Printf("Cap is %d, Len is %d\n", cap(purchase.items), len(purchase.items))
	for i, v := range purchase.items {
		fmt.Printf("%d - %45s $%6.2f      x%d      $%6.2f\n",
			i, v.title, v.itemPrice, v.quantity, v.totalItemPrice)
	}
	fmt.Printf("%78s\n", "=========")
	fmt.Printf("%70s $%6.2f\n", "Item Total", purchase.totalItem)
	fmt.Printf("%70s $%6.2f\n", "Shipping", purchase.shippingCost)
	fmt.Printf("%70s $%6.2f\n", "Order Total", purchase.totalOrder)
	fmt.Printf("---------------------------\n")
}

func readCustomerData(ordersp *[]order, i int, scanner *bufio.Scanner, transOrPurchase string) {
	var addr *address
	var parts []string
	var toString *string

	line := scanner.Text()

	dataSections := 2
	if transOrPurchase == "Transaction" {
		dataSections = 3
	}
	// for loop to record customer data
	for j := 0; j < dataSections; j++ {
		if j%2 == 0 {
			//fmt.Printf("J:%d, Billing\n", j)
			addr = &(*ordersp)[i].billing
		} else {
			//fmt.Printf("J:%d, Shipping\n", j)
			addr = &(*ordersp)[i].shipping
		}
		i := 0
		for exitLoop := false; exitLoop == false && i < 15; {
			i++
			parts = strings.Split(line, ": ")
			scanner.Scan()
			line = scanner.Text()
			parts[0] = strings.TrimSpace(parts[0])

			switch parts[0] {
			case "First Name":
				toString = &addr.firstName
				*toString = parts[1]
			case "Last Name":
				toString = &addr.lastName
				*toString = parts[1]
			case "Address":
				toString = &addr.street
				*toString = parts[1]
			case "City":
				toString = &addr.city
				*toString = parts[1]
			case "Country":
				toString = &addr.country
				*toString = parts[1]
			case "Country (From above Shipping Calc.)":
				toString = &addr.country
				*toString = parts[1]
			case "Billing State":
				toString = &addr.state
				*toString = parts[1]
			case "State":
				toString = &addr.state
				*toString = parts[1]
				if len((*ordersp)[i].billing.state) < 1 {
					(*ordersp)[i].billing.state = parts[1]
				}
			case "Delivery State":
				toString = &addr.state
				*toString = parts[1]
			case "Postal Code":
				toString = &addr.zipCode
				*toString = parts[1]
			case "Email":
				toString = &addr.email
				*toString = parts[1]
				if transOrPurchase == "Purchase" {
					// Type "Purchase Report"" not "Transaction Report"
					exitLoop = true
				}
			case "Phone":
				toString = &addr.phoneNumber
				*toString = parts[1]
				if transOrPurchase == "Purchase" {
					// Type "Purchase Report"" not "Transaction Report"
					exitLoop = true
				}
			case "State/Province":
				toString = &addr.state
				*toString = parts[1]
				if j == 2 {
					exitLoop = true
				}
			case "Shipping Details":
				exitLoop = true
			case "Other Details":
				exitLoop = true
			default:
				from := "Billing"
				if j%2 == 1 {
					from = "Shipping"
				}
				if !strings.Contains(parts[0], ":") {
					// Idea here is to append this line to the last input
					// For example: street address appended with apartment number
					if toString == nil {
						//log.Fatal("readCustomerData()::toString is nil!!!")
						//tmp := ""
						//toString = &tmp
						// may be useful
						// However, if toSting is nil, we haven't got to the data block yet
						// so just ignore it (FYI there is a "Purchase # xxxx" here that
						// may be useful)
					} else {
						if parts[0] != "" {
							*toString += ", " + parts[0] // note index 0 no 1
							fmt.Printf("\n****updated/combined: \"%s\" from %s\n\n", *toString, from)
						}
					}
				} else {
					fmt.Printf("\n****missed: \"%s\" from %s\n\n", parts[0], from)
				}
			}
		}
	} // end for loop to record customer data
}

func collectItemsPurchaseReport(orderp *order, scanner *bufio.Scanner) {
	totalItems := 0

	//Collect the purchased items here (purchase report)
	orderp.items = make([]item, 0) //item_array[totalItems:0]
	runningTotal := float32(0)
	// for loop to load purchase report data
	for exitLoop := false; exitLoop == false; {
		scanner.Scan()
		line := scanner.Text()

		if strings.Contains(line, "Total:") {
			exitLoop = true
			orderp.totalOrder = getPrice(line)
		} else if strings.Contains(line, "-") {

			quant := getQuantity(line)
			title := getTitle(line)
			tip := getPrice(line)
			price := tip / float32(quant)
			a := item{
				quantity:       quant,
				title:          title,
				itemPrice:      price,
				totalItemPrice: tip,
			}
			runningTotal = runningTotal + tip

			orderp.items = append(orderp.items, a)
			totalItems++
		}
	} //end load purchase report data

	orderp.totalItem = runningTotal
	orderp.shippingCost = orderp.totalOrder - runningTotal
}

func getTransTitle(line string) string {
	return strings.TrimSpace(line[0:strings.Index(line, "$")])
}
func getTransPrice(line string) float32 {
	temp := line[strings.Index(line, "$")+1:]
	temp = temp[0:strings.IndexAny(temp, " \t")]
	price, err := strconv.ParseFloat(temp, 32)
	if err != nil {
		fmt.Printf("Price Error: %s", err)
		price = -1
	}
	return float32(price)
}
func getTransQuantity(line string) int {
	temp := line[strings.Index(line, "$"):]
	temp = temp[strings.IndexAny(temp, " \t"):]
	temp = temp[strings.IndexAny(temp, "0123456789"):]
	temp = temp[:strings.IndexAny(temp, " \t")]
	quant, err := strconv.ParseInt(temp, 10, 32)
	if err != nil {
		fmt.Printf("Quant error: %s", err)
		quant = -1
	}
	return int(quant)
}
func getTransTotalPrice(line string) float32 {
	temp := strings.TrimSpace(line[strings.LastIndex(line, "$")+1:])
	price, err := strconv.ParseFloat(temp, 32)
	if err != nil {
		fmt.Printf("TotalPrice Error: %s", err)
		price = -1
	}
	return float32(price)
}

func collectItemsTransactionReport(orderp *order, scanner *bufio.Scanner) {
	totalItems := 0

	//Collect the purchased items here (transaction report)
	orderp.items = make([]item, 0) //item_array[totalItems:0]
	runningTotal := float32(0)
	// for loop to load purchase report data
	for exitLoop := false; exitLoop == false; {
		scanner.Scan()
		line := scanner.Text()

		if strings.Contains(line, "Total:") {
			orderp.totalOrder = getPrice(line)
		} else if strings.Contains(line, "Name	 Price	 Quantity	 Item Total") {
			// Entering the item block
			for {
				scanner.Scan()
				line := scanner.Text()
				if strings.Index(line, "$") < 0 {
					// Check for "$" is just to see if there is an item in this line
					//fmt.Printf("### No $ fould exit loop\n")
					exitLoop = true
					break
				}
				quant := getTransQuantity(line)
				title := getTransTitle(line)
				tip := getTransTotalPrice(line)
				price := getTransPrice(line)
				a := item{
					quantity:       quant,
					title:          title,
					itemPrice:      price,
					totalItemPrice: tip,
				}
				runningTotal = runningTotal + tip

				orderp.items = append(orderp.items, a)
				totalItems++
			}
		}
	} //end load purchase report data

	orderp.totalItem = runningTotal
	orderp.shippingCost = orderp.totalOrder - runningTotal
}

func createPackingSlip(curOrder order, xlsx string) {
	var input string
	fmt.Printf("Print %s's packing slip? (Y/N/S)\n", curOrder.billing.firstName)
	fmt.Scanf("%s\n", &input)
	for {
		if input == "S" || input == "s" {
			outputSpreadsheet(curOrder, xlsx, true)
			break
		} else if input == "Y" || input == "y" {
			outputSpreadsheet(curOrder, xlsx, false)
			break
		} else if input == "N" || input == "n" {
			break
		} else {
			fmt.Printf("Need a Y, N or S\n")
			fmt.Scanf("%s\n", &input)
		}
	}
}

// main processes customer input csv file, displaying the orders and printer a packing list
func main() {

	//fpath := "C:\\home\\auld\\goDev\\src\\packingSlips\\packingSlipTemplate.xlsx"
	defaultXlsx := "packingSlipTemplate.xlsx"
	versionPtr := pflag.Bool("version", false, "program version")
	xlsxPtr := pflag.String("xlsx", defaultXlsx, "xlsx template file")
	pathPtr := pflag.String("input", "orders.csv", "input customer file in csv format")
	listItPtr := pflag.Bool("listIt", false, "list the store items")

	pflag.Parse()
	//fmt.Println("input:", *pathPtr)
	//fmt.Println("tail:", pflag.Args())

	if *versionPtr == true {
		fmt.Printf("\t Version %d.%d", version.major, version.minor)
		os.Exit(0)
	}
	if *listItPtr == true {
		for k, v := range storeItems {
			fmt.Printf("key[%s] value[%s]\n", k, v)
		}
		os.Exit(0)
	}

	var orders = make([]order, 100)

	//
	// Open Customer data file
	//
	f, err := os.Open(*pathPtr)
	if err != nil {
		fmt.Printf("could not open CVS file: %s\n", *pathPtr)
		os.Exit(1)
		//panic(e)
	}
	//fmt.Printf("default[%s], xlsxPtr[%s]\n", defaultXlsx, *xlsxPtr)
	if defaultXlsx == *xlsxPtr {
		cwd, _ := os.Getwd()
		j := cwd + "\\" + *xlsxPtr
		xlsxPtr = &j
	}

	scanner := bufio.NewScanner(f)
	i := 0
	for scanner.Scan() {
		line := scanner.Text()
		if strings.Contains(line, "Purchase Report") {
			readCustomerData(&orders, i, scanner, "Purchase")
			collectItemsPurchaseReport(&orders[i], scanner)
			printPurchaseRecord(orders[i], i+1)
			createPackingSlip(orders[i], *xlsxPtr)
			i++
		} else if strings.Contains(line, "Transaction Report") {
			collectItemsTransactionReport(&orders[i], scanner)
			readCustomerData(&orders, i, scanner, "Transaction")
			printPurchaseRecord(orders[i], i+1)
			createPackingSlip(orders[i], *xlsxPtr)
			i++
		}
	}
}
