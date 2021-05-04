<!DOCTYPE html>
<html>
<head>
	<title>Login Page</title>
	<!-- Latest compiled and minified CSS -->
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">

	<!-- Optional theme -->
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">

	<!-- Latest compiled and minified JavaScript -->
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
</head>
<body>

    <div class="container">

        <div class="panel panel-info">
            <?php if (!$isLoggedIn){?>
            <div class="panel-heading">Login Form</div>
            <div class="panel-body">

                <button onclick="loginWithOffice365()" type="button" class="btn btn-success">Login with Office365</button>

            </div>
            <?php } else { ?>
                <h3>Welcome <b><?php echo session()->get('user_name');?></b></h3>
                <a class="btn btn-danger" href="<?php echo base_url('home/logout')?>">Log Out</a>
            <?php } ?>
        </div>



    </div>

</body>


<!-- Bootstrap/jQuery -->

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js"></script>

<!-- Moment.js -->
<script src="https://cdn.jsdelivr.net/npm/moment@2.27.0/moment.min.js"></script>
<script src="https://momentjs.com/downloads/moment-timezone-with-data-10-year-range.js"></script>

<!-- MSAL -->
<script src="https://alcdn.msauth.net/browser/2.1.0/js/msal-browser.min.js"></script>

<!-- Graph SDK -->
<script src="https://cdn.jsdelivr.net/npm/@microsoft/microsoft-graph-client/lib/graph-js-sdk.js"></script>

<script src="<?php echo base_url('')?>/assets/js/msauth.js"></script>

</html>