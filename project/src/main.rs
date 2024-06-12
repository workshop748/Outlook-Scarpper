use std::process;


fn main() {
   let output = process::Command::new("python3").arg("theEnd.py").spawn();
}
