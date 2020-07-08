<!doctype html>
<html lang="pts">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">

    <link rel="stylesheet" type="text/css" href="estilo.css" media="screen" />

    <title>PDF para TXT</title>
  </head>
  <body>

<div class="container Div">

  <h2></h2><br>

  <div class="form-group">
    <select class="form-control" id="Combo">
      <option value="MODERNIZACAO" selected>MODERNIZAÇÃO</option>
      <option value="EXPANSAO">EXPANSÃO</option>
      <option value="NP">NP</option>
    </select>
  </div>

  <div class="form-group"> 
    <div class="custom-file">
      <input type="file" class="custom-file-input" id="fileUpload" name="arquivo">
      <label class="custom-file-label" for="arquivo" >Selecionar arquivo EXCEL</label>
    </div>
  </div>

  <div class="form-group">
    <input type="button" id="upload" value="Gerar Arquivo TXT" class="btn btn-primary">
  </div>
</div>


<hrs/>
<div id="dvExcel"></div>
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.13.5/xlsx.full.min.js"></script>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.13.5/jszip.js"></script>
<script type="text/javascript" src="node_modules/file-saver/FileSaver.min.js"></script>


<script type="text/javascript">

    $("body").on("click", "#upload", function () {

        var fileUpload = $("#fileUpload")[0];
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;

        if (regex.test(fileUpload.value.toLowerCase())) {

            if (typeof (FileReader) != "undefined") {

                var reader = new FileReader();
 
                if (reader.readAsBinaryString) {

                    reader.onload = function (e) {
                        ProcessExcel(e.target.result);
                    };
                    reader.readAsBinaryString(fileUpload.files[0]);

                } else {

                    reader.onload = function (e) {
                        var data = "";
                        var bytes = new Uint8Array(e.target.result);
                        for (var i = 0; i < bytes.byteLength; i++) {
                            data += String.fromCharCode(bytes[i]);
                        }
                        ProcessExcel(data);
                    };
                    reader.readAsArrayBuffer(fileUpload.files[0]);
                }
            } else {
                alert("Esse navegador não suporta HTML5.");
            }
        } else {
            alert("Por favor selecione um arquivo EXCEL.");
        }
    });

    function ProcessExcel(data) {

        var itemSelecionado = $('#Combo').val();

        var workbook = XLSX.read(data, {
            type: 'binary'
        });
 
        // tab 1 e tab 3
        var firstSheet = workbook.SheetNames[0];
 
        // pegar todos os valores das tabs 3 
        var excelRowsFirstSheet = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);

        var NumOrcamento  = excelRowsFirstSheet[1].__EMPTY_1;
        var DataOrcamento = excelRowsFirstSheet[2].__EMPTY_7.replace(" Dias","");
        var Texto = '';

        for (var i = 10; i < excelRowsFirstSheet.length; i++) {

          if(excelRowsFirstSheet[i].__EMPTY_5 != undefined){
            Texto = Texto +  NumOrcamento + "," + DataOrcamento + ",17." + excelRowsFirstSheet[i].__EMPTY + "," 
                          + excelRowsFirstSheet[i].__EMPTY_5 + ",1.000," + itemSelecionado + ",\n";  
          
          }

        }

        var blob = new Blob([Texto], {type: "text/plain;charset=utf-8"});
          saveAs(blob, NumOrcamento);
    }
 
</script>


    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
</html>