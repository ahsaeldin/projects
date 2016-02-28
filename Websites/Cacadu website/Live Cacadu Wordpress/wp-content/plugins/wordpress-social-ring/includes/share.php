<!DOCTYPE html>
<html>
	<head>
		<style type="text/css">
			body {
			  color: #333333;
			  direction: ltr;
			  font-family: "lucida grande",tahoma,verdana,arial,sans-serif;
			  font-size: 11px;
			  line-height: 1.28;
			  margin: 0;
			  padding: 0;
			  text-align: left;
			}

			.share-container {
				display: inline-block;
				margin: 0px;
				padding: 0px;
				background: transparent;
				border-style: none;
				float: none;
				border-spacing: 0;
				font-size: 11px;
				line-height: normal;
				font-family: "lucida grande",tahoma,verdana,arial,sans-serif !important;
				line-height: 15px;
			}
			
			.facebook-share-button {
				background: url("../images/facebook.png") no-repeat scroll 4px 3px #ECEEF5 ;
				border: 1px solid #CAD4E7;
				border-radius: 3px 3px 3px 3px;
				color: #3B5998 !important;
				padding: 2px 4px 2px 22px !important;
				text-decoration: none;	
				display: block;
				line-height: 14px !important;
			}
		
		</style>
		<script type="text/javascript">
			av_toolbar_off=1;
		</script>
	</head>
	<body marginwidth="0" marginheight="0">
		<table style="border-collapse:collapse;">
			<tbody>
				<tr>
					<td style="border-spacing: 0;font-size: 11px;line-height: normal;padding: 0;">
						<div class="share-container">
							<a onclick="window.open('http://www.facebook.com/sharer.php?u=<?php echo $_GET['url']; ?>', 'facebook', 'width=640,height=480');return false" class="facebook-share-button" href="#">Share</a>
						</div>
					</td>
				</tr>
			</tbody>
		</table>
	</body>
</html>
