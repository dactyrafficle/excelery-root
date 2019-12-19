
let changeImagesBtn = document.getElementById('changeImagesBtn');

chrome.storage.sync.get('color', function(data) {
	//changeImagesBtn.style.backgroundColor = data.color;
	//changeImagesBtn.setAttribute('value', data.color);
});
	

// just like in a webpage, im adding an event listener
changeImagesBtn.addEventListener('click', function(e) {

	//let color = e.target.value;
	//console.log('ok you clicked ' + Math.floor(Math.random()*899+100));

	chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
		chrome.tabs.executeScript(
			tabs[0].id, // you must indicate the tab being accessed to get to dom
			{file: "changeImages.js"} // include the code file in the manifest
		); 
	});

});


	
	