      /IF NOT DEFINED(DEF_NEW_STRING)
       DCL-PR String$new Object(*JAVA:stringClass)
           EXTPROC(*JAVA:stringClass:*CONSTRUCTOR);
         value VARCHAR(65535) Const;
       END-PR;
       DCL-C stringClass 'java.lang.String';
      /DEFINE DEF_NEW_STRING
      /ENDIF
       DCL-PR XTool$new Object(*JAVA:xlClass)
         EXTPROC(*JAVA:xlClass:*CONSTRUCTOR);
         fileName VARCHAR(255) Const;
       END-PR;
       DCL-PR XTool$setRow EXTPROC(*JAVA:xlClass:'setRowValues');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
         values Object(*JAVA:'java.lang.Object') dim(18) CONST;
       END-PR;

       DCL-PR copyRow EXTPROC(*JAVA:xlClass:'copyRow');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
       END-PR;

       DCL-PR setCellBytes EXTPROC(*JAVA:xlClass:'setCellBytes');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
         dvalue VARCHAR(65535) CONST;
       END-PR;
       DCL-PR setCellExcelDate EXTPROC(*JAVA:xlClass:'setCellExcelDate');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
         dvalue VARCHAR(65535) CONST;
       END-PR;
       DCL-PR deleteRow EXTPROC(*JAVA:xlClass:'deleteRow');
         rowNum INT(10) VALUE;
       END-PR;
       DCL-PR getCellBytes VARCHAR(65535)
         EXTPROC(*JAVA:xlClass:'getCellBytes');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
       END-PR;

       DCL-PR setCellInt EXTPROC(*JAVA:xlClass:'setCellInt');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
         dvalue INT(10) VALUE;
       END-PR;
       DCL-PR getCellInt INT(10)
         EXTPROC(*JAVA:xlClass:'getCellInt');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
       END-PR;

       DCL-PR setSheet EXTPROC(*JAVA:xlClass:'setSheet');
         sheetNum INT(10) VALUE;
       END-PR;

       DCL-PR setSheetName EXTPROC(*JAVA:xlClass:'setCurrentSheetName');
         sheetName CHAR(10) CONST;
       END-PR;

       DCL-PR setCellStr EXTPROC(*JAVA:xlClass:'setCellString');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
         dvalue OBJECT(*JAVA:STRINGCLASS) CONST;
       END-PR;
       DCL-PR setCellFormula EXTPROC(*JAVA:xlClass:'setCellFormula');
         rowNum INT(10) VALUE;
         colNum INT(10) VALUE;
         dvalue VARCHAR(255) CONST;
       END-PR;

       DCL-PR XTool$save EXTPROC(*JAVA:xlClass:'save');
         filename VARCHAR(255) Const;
       END-PR;
       DCL-PR XTool$recalc EXTPROC(*JAVA:xlClass:'recalc');
       END-PR;

       DCL-C xlClass 'com.github.j4ile.poi.ExcelTool';

