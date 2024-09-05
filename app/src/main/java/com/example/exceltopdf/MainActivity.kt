package com.example.exceltopdf

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.tooling.preview.Preview
import com.example.exceltopdf.ui.theme.ExcelToPDFTheme
import java.io.File

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {
            ExcelToPDFTheme {
                Scaffold(modifier = Modifier.fillMaxSize()) { innerPadding ->
                    Greeting(
                        name = "Android",
                        modifier = Modifier.padding(innerPadding)
                    )
                }
            }
        }
    }
}

@Composable
fun Greeting(name: String, modifier: Modifier = Modifier) {
    val context = LocalContext.current
    Text(
        text = "Convert excel to pdf",
        modifier = modifier.clickable {
            val assetManager = context.assets
            val inputStream = assetManager.open("test.xls")
            val pdfFile = File(context.getExternalFilesDir(null), "test.pdf")

            ExcelToPDFUtils.convertExcelToPDF(
                inputStream,
                pdfFile
            )
        }
    )
}

@Preview(showBackground = true)
@Composable
fun GreetingPreview() {
    ExcelToPDFTheme {
        Greeting("Android")
    }
}