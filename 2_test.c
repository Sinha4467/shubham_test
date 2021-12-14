#include <stdio.h>
#include <string.h>
#include <math.h>

int main()
{
	int a,b,c=a+b,d=a-b;
    float e,f,g=e+f,h=e-f;
    
    scanf("%d%d\n", &a,&b);
    scanf("%f%f\n", &e,&f);
    
    printf("%d",c);
    printf("%d\n",d);
    printf("%f",g);
    printf("%f\n",h);
    return 0;
}