//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include <ComObj.hpp>
#include "Unit1.h"
#include <conio.h>
#include <iostream>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
Variant exl;
AnsiString s;
AnsiString arrDay[7] = {"Понеділок", "Вівторок", "Середа", "Четвер", "П'ятниця", "Субота", "Неділя"};
AnsiString arrMonth[12] = {"Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень"};
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{




   if( OpenDialog1->Execute() )
 {
  s = OpenDialog1->FileName;
  exl = CreateOleObject("Excel.Application");
  exl.OlePropertyGet("Workbooks").OleProcedure("Open", s.c_str());
        for (int i = 0; i < 12; i++)
        for (int j = 0; j < 7; j++)
                arr[i][j] = exl.OlePropertyGet("Cells", i + 2, j + 2);
                exl.OlePropertyGet("WorkBooks", 1).OleProcedure("Close", false);
                exl.OleProcedure("Quit");
                ComboBox1->Enabled = true;
                RadioGroup1->Enabled = true;
                Graph(0, 0);
                Chart1->Visible = true;
                }
}
//---------------------------------------------------------------------------






void __fastcall TForm1::N2Click(TObject *Sender)
{
        Button1->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N3Click(TObject *Sender)
{
        Close();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N4Click(TObject *Sender)
{
Application->MessageBox("V 1.0\n Автор Гайдучик Дмитро", "Про програму", MB_OK);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
Close();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::RadioGroup1Click(TObject *Sender)
{
switch (RadioGroup1->ItemIndex)
{
 case 0:
 {
  ComboBox1->Items->Clear();
  ComboBox1->Items->Add("По місяцях");
  ComboBox1->Items->Add("По днях");
  ComboBox1->ItemIndex = 0;
  break;
 }
 case 1:
 {
  ComboBox1->Items->Clear();
  for (int i = 0; i < 12; i++)
  ComboBox1->Items->Add(arrMonth[i]);
  ComboBox1->ItemIndex = 0;
  break;
 }
 case 2:
 {
  ComboBox1->Items->Clear();
    for (int i = 0; i < 7; i++)
  ComboBox1->Items->Add(arrDay[i]);
  ComboBox1->ItemIndex = 0;
  break;
 }
}
Graph(RadioGroup1->ItemIndex, ComboBox1->ItemIndex);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox1Change(TObject *Sender)
{
Graph(RadioGroup1->ItemIndex, ComboBox1->ItemIndex);
}
//---------------------------------------------------------------------------

int TForm1::sum(int j, bool month)
{
int sum = 0;
if (month)
{
 for (int i = 0; i < 7; i++)
 {
  sum += arr[j][i];
 }
 }
 else
 {
  for (int i = 0; i < 12; i++)
 {
  sum += arr[i][j];
 }
 }
 return sum;
}

void TForm1::Graph(int radio, int combo)
{
Series1->Clear();
 switch (radio)
 {
  case 0:
  {
   if (combo == 0)
   {
        Chart1->BottomAxis->Maximum = 11;
  for (int i = 0; i < 12; i++)
   Series1->AddXY(i, sum(i, true), arrMonth[i], clRed);
   }
   else
   {
    Chart1->BottomAxis->Maximum = 6;
  for (int i = 0; i < 7; i++)
   Series1->AddXY(i, sum(i, false), arrDay[i], clRed);
   }
   break;
  }
  case 1:
  {
  Chart1->BottomAxis->Maximum = 6;
  for (int i = 0; i < 7; i++)
   Series1->AddXY(i, arr[combo][i], arrDay[i], clRed);
   break;
  }
  case 2:
  {
    Chart1->BottomAxis->Maximum = 11;
  for (int i = 0; i < 12; i++)
   Series1->AddXY(i, arr[combo][i], arrMonth[i], clRed);
   break;
  }

 }
}

