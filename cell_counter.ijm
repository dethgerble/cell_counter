// ******** MACRO - Main Image processing macro ******** 
macro "Setup [F5]" {
	// anlyzes a folder containing .jpg images of cells and counts them

	
   	directory_path = getDirectory("Choose the folder with images of cells to run cell counter on (must be .jpg)");
	

	input_directory = replace(directory_path, "//", "////");
	process_folder(input_directory);

}	


// ******* MACRO - End macro **********************




function make_directory(input_folder) {
	// Takes in input folder and creates new results directory, if it doesn't exist
	dir = input_folder; 
	new_directory = dir + "results"; 

	// Create directory if nonexistent
	if (!File.isDirectory(new_directory)){
		File.makeDirectory(new_directory); 
	}
	// return directory name
	return new_directory + "/";
}

function process_folder(input_folder) {
	// Processes the folder
	setBatchMode(true);  // makes process repeatable

	files = getFileList(input_folder);
	
	total_cells = 0;
	
	for (i=0; i < files.length; i++) {
		// ** Ignore subdirectories **
		//if (File.isDirectory(input_folder + files[i]))
		//process_folder("" + input_folder + files[i]);
		

		if(endsWith(files[i], ".jpg")) {
			// Count cells if is a .jpg, then update total_cells
			num_cells = process_file(files[i], input_folder);
			total_cells = total_cells + num_cells;


			if (num_cells == 0) {
			// Cell counting error
				title = "No Results !! ";
		  		msg = "Please adjust parameters \n and retry analysis.";
				Dialog.createNonBlocking(title);
				Dialog.addMessage(msg); 
				Dialog.setLocation(10, 10);
				Dialog.show()
				break;
			}
		}

	}
	setBatchMode(false);
	// end batch process


	// display results
	print("Total number of cells: " + total_cells);
	
	// Close open windows
	while (nImages>0) { 
    		selectImage(nImages); 
   		close(); 
	} 
}

function process_file(file_name, folder) {
	// Close open windows
      	while (nImages>0) { 
          		selectImage(nImages); 
         		close(); 
      	} 

  // Close ROI Manager
      	if (isOpen("ROI Manager")) {
     		selectWindow("ROI Manager");
     		run("Close");
  	}

  // Close Results Window
      	if (isOpen("Results")) {
     		selectWindow("Results");
     		run("Close");
  	}

  // Close Threshold Window
      	if (isOpen("Threshold")) {
     		selectWindow("Threshold");
     		run("Close");
  	}

	//Do processing here

	lower_particle_size 		= 0.000005;
	upper_particle_size		= "Infinity";
	upper_threshold 			= 215;
	erode_dilate_cycles 		= 3;
	image_scale			= 647;


	// open image
	image_name_path = file_name;
	open(folder + image_name_path);
	
	image_directory_path = folder;

	// Make output image directory path
	output_directory_path = make_directory(image_directory_path);

	// file processing
	image_name 			= File.getName(image_name_path);
	image_name_no_file_ext 	= replace(image_name, ".jpeg","");
	image_name_no_file_ext 	= replace(image_name, ".jpg","");
	composite_name 			= image_name_no_file_ext +"_composite.jpg";
	starting_image			= image_name_no_file_ext +".jpg";
	  	

	 // Set image scale *** This must be set for each study.*** Each image should be at the same magnification to ensure the scale is correct	
	run("Set Scale...", "distance=image_scale known=1 unit=mm global");  // Scale 


	// Duplicate images for analysis	
	originalName 	= getTitle(); 
	original_Id 		= getImageID(); 
	run("Duplicate...", "title=starting_image");

	// Invert is very very important if background is black
  
	run("Invert");

	// Turn to 8-bit (B&W)
	run("8-bit");

	// PREPARE IMAGE

	run("Subtract Background...", "rolling=50 light");


	// *** THRESHOLD CELLS***

	setAutoThreshold("Default");
	setThreshold(0, 215);
	run("Convert to Mask");
	// image is binary, now can manipulate


	// Close Threshold Window
	if (isOpen("Threshold")) {
	  	selectWindow("Threshold");
	   	run("Close");
	}


	// Clean up dirt on image
	run("Fill Holes");
	

	// loop though image processing to fix issues where cells are detected as multiple cells
	dilate = erode_dilate_cycles + 2;
	erode = erode_dilate_cycles;

	for (i=0; i<erode; i++) {
		run("Erode");
	}

	for (i=0; i<dilate; i++) {
		run("Dilate");
		run("Close-");
	}
		
	
	run("Despeckle");

	// Process image and anlyze the number of cells
	
	run("Watershed");
	run("Set Measurements...", "area perimeter shape feret's limit display decimal=7");
	run("Analyze Particles...", "size=lower_particle_size-upper_particle_size show=Outlines display clear add");

	// record number of cells

	cell_num = nResults;
	run("Summarize");

	// rename and close original/manipulated duplicate
	selectImage("starting_image"); 
	roiManager("Show All");
	roiManager("Set Line Width", 3);
	run("Flatten");
	rename("composite_image");
	selectWindow("starting_image");
   	run("Close");

	selectWindow(originalName);
	setLocation(20, 50); 
	run("Close");
	selectWindow("composite_image");
	setLocation(800, 50); 

	// Save processed image 

	selectImage("composite_image");
	save(output_directory_path + composite_name);

	// Save Measurement Results	

	exl_file_name = image_name_no_file_ext + "_Results" + ".xls"; 
	saveAs("Results", output_directory_path + exl_file_name);
	processed = true;		

	
	// Close results if open
	if (isOpen("Results")) {
			selectWindow("Results");
   		run("Close");
	}

	// Close the final picture
	close();

	// Return counted cells
	return cell_num;
}