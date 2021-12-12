#include<stdio.h>
int main()
{
    int a,b,c,d;
    printf("Enter any four numbers\n");
    scanf("%d%d%d%d",&a,&b,&c,&d);
    if(a>b)
    {
        if(a>c){
            if(a>d){
                printf("The number greatest is:%d",a);
            }
            else{
                printf("The number greatest is:%d",d);
            }
        }
        else
        {
            printf("The number greatest is:%d",c);
        }
    }
    else if(b>c){
        if(b>d){
            printf("The number greatest is:%d",b);
        }
        else{
            printf("The number greatest is:%d",d);
        }
    }
    else if (c>d){
        printf("The number greatest is:%d",c);
    }
    else{
        printf("The number greatest is:%d",d);
    }
     return 0;
}