from MyQR import myqr
import os

# Open the students.txt file and read the lines
with open('students.txt', 'r') as f:
    lines = f.read().strip().split("\n")

# Iterate through each line to generate a QR code
for line in lines:
    if not line:
        continue  # Skip empty lines

    data, student_name = line.split(',')  # Split each line into barcode and student name
    qr_data = f"{data} - {student_name}"  # Format the data for the QR code
    
    # Generate the QR code
    version, level, qr_name = myqr.run(
        str(qr_data),
        level='H',
        version=1,
        picture="Bg.png",  # Optional: Remove or change if you don't need a background image
        colorized=True,
        contrast=1.0,
        brightness=1.0,
        save_name=f"{data}_{student_name}.png",  # Save the QR code with a meaningful name
        save_dir=os.getcwd()
    )
