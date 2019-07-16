doctype = """
<!DOCTYPE html>
<html>
 <head>
  <meta charset="UTF-8">
  <title>{file} - {sheet} - {range_start}:{range_end}</title>
  <style type="text/css">
   .plus {{
     background-image: url("https://gurufocus.s3.amazonaws.com/images/icons/gray/png/sq_plus_icon&16.png");
     width: 16px;
     padding: 0;
     background-repeat: no-repeat;
   }}
   .hidden {{
    display: none;
   }}
   .bold {{
    font-weight: bold;
   }}
   .gray {{
    color: #333;
   }}
   table {{
    width: 100%;
    border-collapse: collapse;
    border-bottom: 1px solid black;
    text-align: center;
    margin-bottom: 2em;
   }}
   body > table {{
    margin-top: 2em;
   }}
   body > table > tbody {{
    border-right: 1px solid black;
   }}
   td:first-child {{
    border-left: 1px solid gray;
   }}
   td {{
    border-top: 1px solid black;
   }}
   #stats {{
    text-align: center;
    font-weight: bold;
    margin: 1em;
   }}
  </style>
  <script type="text/javascript">
   function toggleTable(elm) {{
    var table = elm.parentElement.nextElementSibling;
    if(table.className == 'hidden')
      table.className = '';
    else
      table.className = 'hidden';
   }}
  </script>
 </head>
 <body>
  <h1>LMSD lookup report</h1>
  <h2>{file} - {sheet} - {range_start}:{range_end}</h2>
  Search parameters: mass for adducts {adducts} &plusmn; {error} {error_unit} - highlighted masses: intensity >= {threshold}"""

mass_start = """
  <table>
   <thead>
    <tr>
     <th></th>
     <th>Mass</th>
     <th>Intensity</th>
     <th>Results</th>
    </tr>
   </thead>
   <tbody>"""

mass = """
    <tr>
     <td class="plus" onclick="toggleTable(this)"></td>
     <td class="{bold}">{mass}</td>
     <td class="{bold}">{intensity}</td>
     <td>{total} results (in {isomers} isomer groups)</td>
    </tr>"""

formula_start = """
    <tr class="hidden">
     <td></td>
     <td colspan="3">
      <table>
       <thead>
        <tr>
         <th></th>
         <th>Formula</th>
         <th>Mass</th>
         <th>Adduct type</th>
         <th colspan="2">Error</th>
         <th>Results</th>
        </tr>
       </thead>
       <tbody>"""

formula = """
        <tr>
         <td class="plus" onclick="toggleTable(this)"></td>
         <td>{formula}</td>
         <td>{mass}</td>
         <td>M+{adduct}</td>
         <td align="right">{error_abs} Da</td>
         <td align="left">{error_rel} ppm</td>
         <td>{count}</td>
        </tr>"""

entry_start = """
        <tr class="hidden">
         <td></td>
         <td colspan="5">
          <table>
           <thead>
            <tr>
             <th>LM_ID</th>
             <th>Common Name</th>
             <th>Systematic Name</th>
             <th>Category</th>
             <th>Main Class</th>
             <th>Sub Class</th>
            </tr>
           </thead>
           <tbody>"""

entry = """
            <tr>
             <td>
              <a target="_blank"
               href="http://www.lipidmaps.org/data/LMSDRecord.php?LMID={LM_ID}">
                {LM_ID}
              </a>
             </td>
             <td>{COMMON_NAME}</td>
             <td>{SYSTEMATIC_NAME}</td>
             <td>{CATEGORY}</td>
             <td>{MAIN_CLASS}</td>
             <td>{SUB_CLASS}</td>
            </tr>"""

entry_end = """
           </tbody>
          </table>
         </td>
        </tr>"""

formula_end = """
       </tbody>
      </table>
     </td>
    </tr>"""

mass_end = """
   </tbody>
  </table>"""

end = """
  <div id="stats">
   Found matches for {matches} of the {masses}
   selected masses ({total} total molecules)
  </div>
 </body>
</html>"""
