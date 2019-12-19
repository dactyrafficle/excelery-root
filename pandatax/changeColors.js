let all = document.getElementsByTagName('*');

for (let i = 0; i < all.length; i++) {
	let el = all[i];

	try {
		el.style.color = '#fcf';
		el.style.backgroundColor = "#333";
	}
	catch(err) {
		console.log(Math.random());
	}
}