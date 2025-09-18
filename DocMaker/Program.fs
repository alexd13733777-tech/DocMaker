open System
open System.Windows.Forms
open Microsoft.Office.Interop.Excel
open System.IO
open System.Drawing
[<STAThread>]
let mutable widthL: double list = []
let mutable heightL: double list = []
let mutable countL: double list = []
let mutable orderL: string list = []
let mutable rowStartPosition = 0
let mutable columnStartPosition = 0

let valueByKeySearchD (key: string) (text: string) =
    let words = text.Split(' ')
    let mutable found = false
    let results = ResizeArray<float>()
    for word in words do
        if found then
            match System.Double.TryParse(word) with
            | true, value -> results.Add(value)
            | _ -> found <- false
        elif word = key then
            found <- true
    results |> List.ofSeq

let valueByKeyCopy (keySynonyms: ResizeArray<string>) (worksheet: Worksheet) =
    let results = ResizeArray<string>()
    for i in 1 .. 29 do
        for j in 1 .. 29 do
            let cell = worksheet.Cells.[i, j] :?> Range
            let value = Convert.ToString(cell.Value2)
            if not (isNull value) then
                for key in keySynonyms do
                    if value.Contains(key) then
                        // Сканируем вниз по столбцу
                        for k in i + 1 .. i + 3 do
                            let nextCell = worksheet.Cells.[k, j] :?> Range
                            let nextValue = Convert.ToString(nextCell.Value2)
                            if not (isNull nextValue) then
                                results.Add(nextValue)
    results |> List.ofSeq

let valueByKeyCopyD (keySynonyms: ResizeArray<string>) (worksheet: Worksheet) =
    let results = ResizeArray<float>()
    for i in 1 .. 29 do
        for j in 1 .. 29 do
            let cell = worksheet.Cells.[i, j] :?> Range
            let value = Convert.ToString(cell.Value2)
            if not (isNull value) then
                for key in keySynonyms do
                    if value.Contains(key) then
                        // Сканируем вниз по столбцу
                        for k in i + 1 .. i + 3 do
                            let nextCell = worksheet.Cells.[k, j] :?> Range
                            let nextValue = Convert.ToString(nextCell.Value2)
                            match Double.TryParse(nextValue) with
                            | true, num -> results.Add(num)
                            | _ -> ()
    results |> List.ofSeq


let tableFill () =
    let excelApp = new ApplicationClass()
    let workbook = excelApp.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), "pattern.xls"))
    let worksheet = workbook.Sheets.[1] :?> Worksheet
    for i in 0 .. min 3 (heightL.Length - 1) do
        worksheet.Cells.[i + 6, 1] <- (i + 1).ToString()
        worksheet.Cells.[i + 6, 2] <- orderL.[0]
        worksheet.Cells.[i + 6, 3] <- widthL.[i]
        worksheet.Cells.[i + 6, 4] <- heightL.[i]
        worksheet.Cells.[i + 6, 5] <- countL.[i]
        worksheet.Cells.[i + 6, 7] <- widthL.[i] * heightL.[i] * countL.[i] / 1000000.0
    worksheet.Cells.[11, 2] <- "Конец работы демо версии"
    workbook.SaveCopyAs(Path.Combine(Directory.GetCurrentDirectory(), "SuperDoc777.xls"))
    workbook.Close()
    excelApp.Quit()

let button2_Click (_: obj) (_: EventArgs) =
    let excelApp = new ApplicationClass()
    let openDialog = new OpenFileDialog()
    openDialog.Filter <- "Книга Excel (*.xls)|*.xls|Все файлы (*.xlsx)|*.xlsx"
    if openDialog.ShowDialog() = DialogResult.OK then
        let workbook = excelApp.Workbooks.Open(openDialog.FileName, ReadOnly = false)
        let worksheet = workbook.Worksheets.[1] :?> Worksheet
        // Поиск позиции ключа
        for i in 1 .. 29 do
            for j in 1 .. 29 do
                let cellValue = Convert.ToString((worksheet.Cells.[i, j] :?> Range).Value2)
                if not (isNull cellValue) &&
                   (cellValue.Contains("заказ") || cellValue.Contains("Артикул")) then
                    rowStartPosition <- i
                    columnStartPosition <- j

        // Поиск по ключам
        let keySynonyms = ResizeArray<string>()
        keySynonyms.Add("Высота")
        keySynonyms.Add("Длина")
        heightL <- valueByKeyCopyD keySynonyms worksheet

        keySynonyms.Clear()
        keySynonyms.Add("Кол-во")
        countL <- valueByKeyCopyD keySynonyms worksheet

        keySynonyms.Clear()
        keySynonyms.Add("Артикул")
        keySynonyms.Add("заказ") 
        orderL <- valueByKeyCopy keySynonyms worksheet

        keySynonyms.Clear()
        keySynonyms.Add("Ширина")
        widthL <- valueByKeyCopyD keySynonyms worksheet

        keySynonyms.Clear()
        tableFill()

        MessageBox.Show("Выгрузка завершена успешно", "Результат выгрузки", MessageBoxButtons.OK, MessageBoxIcon.Information) |> ignore
        excelApp.Workbooks.Close()


[<STAThread>]
[<EntryPoint>]
let main _ =
    Application.EnableVisualStyles()
    Application.SetCompatibleTextRenderingDefault(false)
    let form = new Form(Text = "Doc Maker", Width = 300, Height = 200)
    let button = new System.Windows.Forms.Button(Text = "Создать документ", Dock = DockStyle.Fill)
    // Load an image and set it to the button
    let image =  Image.FromFile("Excel pic.png") // Replace with your image path
    button.Image <- image
    button.ImageAlign <- ContentAlignment.MiddleCenter // Optional: Align the image
    button.TextAlign <- ContentAlignment.TopCenter
    button.Click.AddHandler(EventHandler(button2_Click))
    form.Controls.Add(button)
    System.Windows.Forms.Application.Run(form)
    0