#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
int main(int argc, char** argv)
{
    FILE* fileptr;
    unsigned char* buffer;
    long filelen;

    fileptr = fopen(argv[1], "rb");  // Open the file in binary mode
    fseek(fileptr, 0, SEEK_END);          // Jump to the end of the file
    filelen = ftell(fileptr);             // Get the current byte offset in the file
    rewind(fileptr);                      // Jump back to the beginning of the file

    buffer = (unsigned char*)malloc(filelen * sizeof(unsigned char)); // Enough memory for the file
    fread(buffer, filelen, 1, fileptr); // Read in the entire file
    fclose(fileptr); // Close the file
    char* shellcode = malloc(300000 * sizeof(char));
    int bufflength = filelen * 6;
    int i;
    for (i = 0; i < filelen; i++)
    {

        sprintf(&shellcode[i * 6], "0x%02X, ", buffer[i]);

    }
    shellcode[bufflength - 2] = '\0';
    printf("%s\n", shellcode);
    printf("buffer length is %d", filelen);
    free(buffer);
    free(shellcode);
}
