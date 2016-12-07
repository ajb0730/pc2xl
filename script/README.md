### Compile For Windows

1. Install ActiveState Perl 5.

2. Run `ppm install PAR::Packer`

3. Run `pp -B -o pc2xl.exe pc2xl.pl`

4. The above command may fail, complaining about `Can't locate loadable object 
   for module IO in @INC`. You need to locate IO.dll (try
   `C:\Perl64\lib\auto\IO\IO.dll`) and copy it to the same directory as 
   pc2xl.pl.

5. Rinse and repeat until you have it compiling.

6. Package `pc2xl.exe` with all of the DLLs you copied and deploy that to the
   target machine.
