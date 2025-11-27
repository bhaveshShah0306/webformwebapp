<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BulkBookingUpload.aspx.cs" Inherits="WebApplication1.BulkBookingUpload" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Hotel Booking Bulk Upload</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.6.0/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" />
    <style>
        body { background: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .container { max-width: 1200px; margin-top: 40px; }
        .upload-card { background: white; border-radius: 10px; box-shadow: 0 2px 15px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 30px; }
        .upload-zone { border: 3px dashed #007bff; border-radius: 10px; padding: 50px; text-align: center; background: #f8f9fa; cursor: pointer; transition: all 0.3s; }
        .upload-zone:hover { border-color: #0056b3; background: #e7f1ff; }
        .upload-zone.dragover { border-color: #28a745; background: #d4edda; }
        .file-info { display: none; margin-top: 20px; }
        .progress { height: 25px; margin-top: 15px; display: none; }
        .validation-results { margin-top: 20px; max-height: 400px; overflow-y: auto; }
        .header-section { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px 10px 0 0; margin: -30px -30px 30px -30px; }
        .btn-download { background: #17a2b8; color: white; }
        .btn-download:hover { background: #138496; color: white; }
        .status-badge { font-size: 0.85rem; padding: 5px 10px; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />
        
        <div class="container">
            <div class="upload-card">
                <div class="header-section">
                    <h2><i class="fas fa-hotel"></i> Hotel Chain Booking Bulk Upload</h2>
                    <p class="mb-0">Upload Excel files with multiple booking records from tourist locations</p>
                </div>

                <!-- Agent Information -->
                <div class="row mb-4">
                    <div class="col-md-6">
                        <label><i class="fas fa-user"></i> Agent ID:</label>
                        <asp:TextBox ID="txtAgentId" runat="server" CssClass="form-control" placeholder="Enter Agent ID" />
                    </div>
                    <div class="col-md-6">
                        <label><i class="fas fa-map-marker-alt"></i> Location:</label>
                        <asp:TextBox ID="txtLocation" runat="server" CssClass="form-control" placeholder="Tourist Location" />
                    </div>
                </div>

                <!-- Download Template Button -->
                <div class="mb-4">
                    <button type="button" id="btnDownloadTemplate" class="btn btn-download">
                        <i class="fas fa-download"></i> Download Excel Template
                    </button>
                    <small class="text-muted ml-3">Download the template to ensure proper formatting</small>
                </div>

                <!-- Upload Zone -->
                <div id="uploadZone" class="upload-zone">
                    <i class="fas fa-cloud-upload-alt fa-3x text-primary mb-3"></i>
                    <h4>Drag & Drop Excel File Here</h4>
                    <p class="text-muted">or click to browse</p>
                    <asp:FileUpload ID="fileUpload" runat="server" style="display: none;" accept=".xlsx,.xls" />
                </div>

                <!-- File Information -->
                <div id="fileInfo" class="file-info">
                    <div class="alert alert-info">
                        <strong><i class="fas fa-file-excel"></i> Selected File:</strong> <span id="fileName"></span>
                        <button type="button" id="btnClear" class="btn btn-sm btn-danger float-right">
                            <i class="fas fa-times"></i> Clear
                        </button>
                    </div>
                </div>

                <!-- Progress Bar -->
                <div class="progress" id="progressBar">
                    <div class="progress-bar progress-bar-striped progress-bar-animated" role="progressbar" style="width: 0%"></div>
                </div>

                <!-- Upload Button -->
                <div class="text-center mt-3">
                    <asp:Button ID="btnUpload" runat="server" Text="Upload & Process Bookings" 
                        CssClass="btn btn-primary btn-lg" OnClientClick="return validateAndUpload();" 
                        OnClick="btnUpload_Click" style="display: none;" />
                </div>

                <!-- Validation Results -->
                <div id="validationResults" class="validation-results"></div>
            </div>
        </div>
    </form>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.6.0/js/bootstrap.bundle.min.js"></script>
    <script>
        $(document).ready(function() {
            // Upload zone click
            $('#uploadZone').click(function() {
                $('#<%= fileUpload.ClientID %>').click();
            });

            // File selection handler
            $('#<%= fileUpload.ClientID %>').change(function() {
                handleFileSelect(this.files);
            });

            // Drag and drop handlers
            $('#uploadZone').on('dragover', function(e) {
                e.preventDefault();
                e.stopPropagation();
                $(this).addClass('dragover');
            });

            $('#uploadZone').on('dragleave', function(e) {
                e.preventDefault();
                e.stopPropagation();
                $(this).removeClass('dragover');
            });

            $('#uploadZone').on('drop', function(e) {
                e.preventDefault();
                e.stopPropagation();
                $(this).removeClass('dragover');
                
                var files = e.originalEvent.dataTransfer.files;
                if (files.length > 0) {
                    $('#<%= fileUpload.ClientID %>')[0].files = files;
                    handleFileSelect(files);
                }
            });

            // Clear button
            $('#btnClear').click(function() {
                clearFileSelection();
            });

            // Download template
            $('#btnDownloadTemplate').click(function () {
                window.location.href = 'DownloadTemplate.ashx';
            });

        });

        function handleFileSelect(files) {
            if (files.length === 0) return;

            var file = files[0];
            var fileName = file.name;
            var fileExt = fileName.split('.').pop().toLowerCase();

            if (fileExt !== 'xlsx' && fileExt !== 'xls') {
                alert('Please select a valid Excel file (.xlsx or .xls)');
                clearFileSelection();
                return;
            }

            $('#fileName').text(fileName);
            $('#fileInfo').fadeIn();
            $('#<%= btnUpload.ClientID %>').fadeIn();
            $('#validationResults').empty();
        }

        function clearFileSelection() {
            $('#<%= fileUpload.ClientID %>').val('');
            $('#fileInfo').fadeOut();
            $('#<%= btnUpload.ClientID %>').fadeOut();
            $('#validationResults').empty();
            $('#progressBar').hide().find('.progress-bar').css('width', '0%');
        }

        function validateAndUpload() {
            var agentId = $('#<%= txtAgentId.ClientID %>').val().trim();
            var location = $('#<%= txtLocation.ClientID %>').val().trim();
            var fileInput = $('#<%= fileUpload.ClientID %>')[0];

            if (!agentId) {
                alert('Please enter Agent ID');
                $('#<%= txtAgentId.ClientID %>').focus();
                return false;
            }

            if (!location) {
                alert('Please enter Location');
                $('#<%= txtLocation.ClientID %>').focus();
                return false;
            }

            if (!fileInput.files || fileInput.files.length === 0) {
                alert('Please select an Excel file');
                return false;
            }

            // Show progress bar
            $('#progressBar').show();
            animateProgress();
            
            return true;
        }

        function animateProgress() {
            var progress = 0;
            var interval = setInterval(function() {
                progress += 5;
                $('.progress-bar').css('width', progress + '%');
                if (progress >= 90) {
                    clearInterval(interval);
                }
            }, 100);
        }        

        function showResults(success, message, details) {
            $('.progress-bar').css('width', '100%');
            setTimeout(function() {
                $('#progressBar').fadeOut();
                
                var alertClass = success ? 'alert-success' : 'alert-danger';
                var icon = success ? 'fa-check-circle' : 'fa-exclamation-circle';
                
                var html = '<div class="alert ' + alertClass + ' mt-3">';
                html += '<h5><i class="fas ' + icon + '"></i> ' + message + '</h5>';
                
                if (details && details.length > 0) {
                    html += '<hr><ul class="mb-0">';
                    details.forEach(function(detail) {
                        html += '<li>' + detail + '</li>';
                    });
                    html += '</ul>';
                }
                
                html += '</div>';
                
                $('#validationResults').html(html);
            }, 500);
        }
    </script>
</body>
</html>