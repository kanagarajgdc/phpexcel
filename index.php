<!DOCTYPE html>
<html>
<head>
	<title>Sugar</title>
	<link rel="stylesheet" type="text/css" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
</head>
<body>
<div class="container">
	<h1>Upload Excel File</h1>
	<form method="POST" action="upload.php" enctype="multipart/form-data">
		<div class="form-group">
			<label>Choose File</label>
			<input type="file" name="excel" class="form-control" />
		</div>
		<div class="form-group">
			<button type="submit" name="upload" class="btn btn-success">Upload</button>
		</div>
	</form>
</div>
</body>
</html>