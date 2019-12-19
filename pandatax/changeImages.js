console.log(Math.random());

let imgURL1 = "";
let imgURL2 = "https://i.imgur.com/CJcN9iQ.png";
let imgURL3 = "";

(function(){
	let imgs = document.getElementsByTagName('source');
	for (let i = 0; i < imgs.length; i++) {
		imgs[i].srcset = imgURL2;
	}	
})();

(function(){
	let imgs = document.getElementsByTagName('img');
	for (let i = 0; i < imgs.length; i++) {
		imgs[i].src = imgURL2;
		imgs[i].srcset = imgURL2;
	}
})();


