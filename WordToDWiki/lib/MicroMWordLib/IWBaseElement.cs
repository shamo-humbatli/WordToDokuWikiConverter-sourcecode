using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MicroMWordLib.WordContentSelection;
namespace MicroMWordLib
{
    public interface IWBaseElement
    {
        WCSelection ContentSelection { get; set;  }
    }
}
