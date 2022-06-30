
    
    function send() {
        alert("before Record Successfuly1");
        var DocumentActionUserID = 2;
        var Activity = "Review";
        var ActivityDate = "29"
        var CreatedBy = CreateGuid();
        var CreatedDate = new Date().toLocaleString();
        var ModifiedBy = createGuid();
        var ModifiedDate = new Date().toLocaleString();
        alert("before Record Successfuly");
        //if (txtid.length != 0 || txtname.length != 0 || txtsalary.length != 0 || txtcity.length != 0) {
        var connection = new ActiveXObject("ADODB.Connection");
        var connectionstring = "jdbc:sqlserver://blueed.database.windows.net:1433;database=BTService;user=BlueedAdmin@blueed;password={your_password_here};encrypt=true;trustServerCertificate=false;hostNameInCertificate=*.database.windows.net;loginTimeout=30;a Source=.;Initial Catalog=EmpDetail;Persist Security Info=True;User ID=sa;Password=****;Provider=SQLOLEDB";
        connection.Open(connectionstring);
        var rs = new ActiveXObject("ADODB.Recordset");
        rs.Open("insert into DocumentActionUserActivityID values('" + DocumentActionUserID + "','" + Activity + "','"
            + ActivityDate + "','" + CreatedBy + "','" + CreatedDate + "','" + ModifiedBy + "','" + ModifiedDate + "')", connection);
        alert("Insert Record Successfuly");

        connection.close();
        //}
        //else {
        //    alert("Please Enter Employee \n Id \n Name \n Salary \n City ");
        //}

    }
    function CreateGuid() {
        function _p8(s) {
            var p = (Math.random().toString(16) + "000000000").substr(2, 8);
            return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
        }
        return _p8() + _p8(true) + _p8(true) + _p8();
    }


}
